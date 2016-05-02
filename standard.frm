VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{7ED47906-67D7-4D60-ABCD-66C3BA9E3452}#1.0#0"; "csmtpctl.ocx"
Object = "{959AC9FE-B2CE-4117-9CE6-56B273C5848F}#1.0#0"; "csmsgctl.ocx"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form FormMAIN 
   Caption         =   "Standard (Compositional Database)"
   ClientHeight    =   7545
   ClientLeft      =   825
   ClientTop       =   1305
   ClientWidth     =   11325
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "STANDARD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7545
   ScaleWidth      =   11325
   Begin SmtpClientCtl.SmtpClient SmtpClient1 
      Left            =   1680
      Top             =   0
      _cx             =   741
      _cy             =   741
   End
   Begin MailMessageCtl.MailMessage MailMessage1 
      Left            =   2280
      Top             =   0
      _cx             =   741
      _cy             =   741
   End
   Begin VB.Frame Frame2 
      Caption         =   "Standard Information"
      ForeColor       =   &H00FF0000&
      Height          =   2655
      Left            =   4920
      TabIndex        =   4
      Top             =   0
      Width           =   5295
      Begin VB.Label LabelTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2880
         TabIndex        =   17
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label LabelCalculated 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label LabelAtomic 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2880
         TabIndex        =   15
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label LabelTotalOxygen 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label LabelExcess 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label LabelZbar 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2880
         TabIndex        =   12
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         Caption         =   "Total Weight %"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3840
         TabIndex        =   11
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         Caption         =   "Calculated Oxygen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         Caption         =   "Atomic Weight"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3840
         TabIndex        =   9
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         Caption         =   "Z - Bar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3840
         TabIndex        =   8
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         Caption         =   "Total Oxygen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         Caption         =   "Excess Oxygen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label LabelStandard 
         BorderStyle     =   1  'Fixed Single
         Height          =   1335
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   5055
      End
   End
   Begin ComctlLib.StatusBar StatusBarAuto 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   7290
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   15849
            Key             =   "status"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Automation status"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Cancel"
            TextSave        =   "Cancel"
            Key             =   "cancel"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Click here to cancel automation"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Pause"
            TextSave        =   "Pause"
            Key             =   "pause"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Click here to pause automation"
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox TextLog 
      Height          =   3615
      Left            =   0
      TabIndex        =   2
      Top             =   2760
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   6376
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"STANDARD.frx":59D8A
   End
   Begin VB.Timer TimerLogWindow 
      Interval        =   500
      Left            =   480
      Top             =   0
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.Frame Frame1 
      Caption         =   "Standards (double-click to see composition data)"
      ClipControls    =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin SysInfoLib.SysInfo SysInfo1 
         Left            =   840
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.ListBox ListAvailableStandards 
         Height          =   2205
         Left            =   120
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Width           =   4455
      End
   End
   Begin VB.Menu menuFile 
      Caption         =   "&File"
      HelpContextID   =   167
      Begin VB.Menu menuFileNew 
         Caption         =   "New"
         HelpContextID   =   168
      End
      Begin VB.Menu menuFileOpen 
         Caption         =   "Open"
         HelpContextID   =   169
      End
      Begin VB.Menu menuFileSaveAs 
         Caption         =   "Save As"
         HelpContextID   =   170
      End
      Begin VB.Menu menuFileClose 
         Caption         =   "Close"
         HelpContextID   =   171
      End
      Begin VB.Menu menuFileSeparator0 
         Caption         =   "-"
      End
      Begin VB.Menu menuFileImport 
         Caption         =   "Import ASCII File"
         HelpContextID   =   599
      End
      Begin VB.Menu menuFileExport 
         Caption         =   "Export ASCII File"
         HelpContextID   =   600
      End
      Begin VB.Menu menuFileImportSingleRowFormat 
         Caption         =   "Import ASCII File (single row format)"
         HelpContextID   =   601
      End
      Begin VB.Menu menuFileExportSingleRowFormat 
         Caption         =   "Export ASCII File (single row format)"
         HelpContextID   =   602
      End
      Begin VB.Menu menuseparator1 
         Caption         =   "-"
      End
      Begin VB.Menu menuFileInputCalcZAFStandardFormatKratios 
         Caption         =   "Input CalcZAF Standard Format Kratios"
         Enabled         =   0   'False
      End
      Begin VB.Menu menuseparator11 
         Caption         =   "-"
      End
      Begin VB.Menu menuFileAMCSD 
         Caption         =   "Create AMCSD.MDB (American Mineralogist Database)"
         HelpContextID   =   729
      End
      Begin VB.Menu menuFileImportStandardsFromCamecaPeakSight 
         Caption         =   "Import Standards From Cameca PeakSight (Sx.mdb)"
         HelpContextID   =   640
      End
      Begin VB.Menu menuFileImportStandardsFromJEOLTextFile 
         Caption         =   "Import Standards From JEOL Text File (created from Perl script)"
         Enabled         =   0   'False
         HelpContextID   =   779
      End
      Begin VB.Menu menuFileSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu menuFileFileInformation 
         Caption         =   "File Information"
         HelpContextID   =   174
         Shortcut        =   ^F
      End
      Begin VB.Menu menuFileSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu menuFilePrintLog 
         Caption         =   "Print Log"
         HelpContextID   =   175
         Shortcut        =   ^P
      End
      Begin VB.Menu menuFilePrintSetup 
         Caption         =   "Print Setup"
         HelpContextID   =   176
      End
      Begin VB.Menu menuFileSeparator4 
         Caption         =   "-"
      End
      Begin VB.Menu menuFileExit 
         Caption         =   "Exit"
         HelpContextID   =   177
      End
   End
   Begin VB.Menu menuEdit 
      Caption         =   "&Edit"
      HelpContextID   =   178
      Begin VB.Menu menuEditCut 
         Caption         =   "Cut"
         HelpContextID   =   179
      End
      Begin VB.Menu menuEditCopy 
         Caption         =   "Copy"
         HelpContextID   =   180
      End
      Begin VB.Menu menuEditPaste 
         Caption         =   "Paste"
         HelpContextID   =   181
      End
      Begin VB.Menu menuEditSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu menuEditSelectAll 
         Caption         =   "Select All"
         HelpContextID   =   182
      End
      Begin VB.Menu menuEditClearAll 
         Caption         =   "Clear All"
         HelpContextID   =   183
      End
   End
   Begin VB.Menu menuStandard 
      Caption         =   "&Standard"
      HelpContextID   =   184
      Begin VB.Menu menuStandardNew 
         Caption         =   "New"
         HelpContextID   =   185
         Shortcut        =   ^N
      End
      Begin VB.Menu menuStandardModify 
         Caption         =   "Modify"
         HelpContextID   =   187
         Shortcut        =   ^M
      End
      Begin VB.Menu menuStandardDuplicate 
         Caption         =   "Duplicate"
         HelpContextID   =   188
      End
      Begin VB.Menu menuStandardSeparator0 
         Caption         =   "-"
      End
      Begin VB.Menu menuStandardDelete 
         Caption         =   "Delete"
         HelpContextID   =   189
      End
      Begin VB.Menu menuStandardDeleteSelected 
         Caption         =   "Delete Selected"
         HelpContextID   =   190
      End
      Begin VB.Menu menuStandardSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu menuStandardListStandardNames 
         Caption         =   "List All Standard Names"
         HelpContextID   =   606
      End
      Begin VB.Menu menuStandardListStandardNamesZbar 
         Caption         =   "List All Standard Names and Average Z"
      End
      Begin VB.Menu menuStandardListElementalStandardNames 
         Caption         =   "List Elemental Standard Names"
         HelpContextID   =   607
      End
      Begin VB.Menu menuStandardListOxideStandardNames 
         Caption         =   "List Oxide Standard Names"
         HelpContextID   =   608
      End
      Begin VB.Menu menuStandardSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu menuStandardListSelectedStandards 
         Caption         =   "List Selected Standards"
         HelpContextID   =   192
      End
      Begin VB.Menu menuStandardListAllStandards 
         Caption         =   "List All Standards"
         HelpContextID   =   193
      End
      Begin VB.Menu menuStandardListElementalStandards 
         Caption         =   "List Elemental Standards"
         HelpContextID   =   609
      End
      Begin VB.Menu menuStandardListOxideStandards 
         Caption         =   "List Oxide Standards"
         HelpContextID   =   610
      End
   End
   Begin VB.Menu menuOptions 
      Caption         =   "O&ptions"
      HelpContextID   =   194
      Begin VB.Menu menuOptionsSearch 
         Caption         =   "Search (for a standard name string)"
         HelpContextID   =   512
      End
      Begin VB.Menu menuOptionsFind 
         Caption         =   "Find (a specific element range in all standards)"
         HelpContextID   =   195
      End
      Begin VB.Menu menuOptionsMatch 
         Caption         =   "Match (a composition with all standards)"
         HelpContextID   =   616
      End
      Begin VB.Menu menuOptionsModalAnalysis 
         Caption         =   "Modal &Analysis (quantitative phase ID)"
         HelpContextID   =   198
      End
      Begin VB.Menu menuOptionsInterferences 
         Caption         =   "Interferences (calculate spectral overlaps)"
         HelpContextID   =   202
      End
   End
   Begin VB.Menu menuXray 
      Caption         =   "&X-Ray"
      HelpContextID   =   203
      Begin VB.Menu menuXrayXrayDatabase 
         Caption         =   "X-Ray Database"
         HelpContextID   =   204
      End
      Begin VB.Menu menuXraySeparator0 
         Caption         =   "-"
      End
      Begin VB.Menu menuXrayEmissionTable 
         Caption         =   "Emission Table (Ka, Kb, La, Lb, Ma, Mb)"
         HelpContextID   =   206
      End
      Begin VB.Menu menuXrayEdgeTable 
         Caption         =   "Edge Table (K, L-I, L-II, L-III, M-I, M-II, M-III, M-IV, M-V)"
         HelpContextID   =   207
      End
      Begin VB.Menu menuXrayFluorescentYieldtable 
         Caption         =   "Fluorescent Yield Table (Ka, Kb, La, Lb, Ma, Mb)"
         HelpContextID   =   208
      End
      Begin VB.Menu menuXraySeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu menuXrayMACTable 
         Caption         =   "MAC Table (Ka, Kb, La, Lb, Ma, Mb)"
         HelpContextID   =   209
      End
   End
   Begin VB.Menu menuAnalytical 
      Caption         =   "&Analytical"
      HelpContextID   =   210
      Begin VB.Menu menuAnalyticalZAFSelections 
         Caption         =   "ZAF, Phi-Rho-Z, Alpha Factor and Calibration Curve Selections"
         HelpContextID   =   212
      End
      Begin VB.Menu menuAnalyticalConditions 
         Caption         =   "Operating Conditions"
         HelpContextID   =   213
      End
      Begin VB.Menu menuAnalyticalEmpiricalMACs 
         Caption         =   "Empirical &MACs"
         HelpContextID   =   211
      End
      Begin VB.Menu menuAnalyticalSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu menuAnalyticalMQOptions 
         Caption         =   "MQ (Monte-Carlo) Calculations"
         HelpContextID   =   841
      End
      Begin VB.Menu menuAnalyticalPENEPMA 
         Caption         =   "PENEPMA (Monte-Carlo) Calculations"
         HelpContextID   =   842
      End
      Begin VB.Menu menuAnalyticalPENFLUOR 
         Caption         =   "PENEPMA (Secondary Fluorescence Profile) Calculations"
         HelpContextID   =   843
      End
   End
   Begin VB.Menu menuOutput 
      Caption         =   "&Output"
      HelpContextID   =   218
      Begin VB.Menu menuOutputLogWindow 
         Caption         =   "Log Window Font"
         HelpContextID   =   219
      End
      Begin VB.Menu menuOutputDebugMode 
         Caption         =   "Debug Mode"
         HelpContextID   =   220
      End
      Begin VB.Menu menuOutputVerboseMode 
         Caption         =   "Verbose Mode"
      End
      Begin VB.Menu menuOutputExtendedFormat 
         Caption         =   "Extended Format"
         HelpContextID   =   221
      End
      Begin VB.Menu menuOutputSaveToDiskLog 
         Caption         =   "Save To Disk Log"
         HelpContextID   =   222
      End
      Begin VB.Menu menuViewDiskLog 
         Caption         =   "View Disk Log"
         HelpContextID   =   223
      End
      Begin VB.Menu menuOutputSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu menuOutputCalculateElectronandXrayRanges 
         Caption         =   "Calculate Electron and X-ray Ranges"
         HelpContextID   =   781
      End
      Begin VB.Menu menuOutputCalculateAlternativeZbars 
         Caption         =   "Calculate Alternative Zbars"
         HelpContextID   =   225
      End
      Begin VB.Menu menuOutputCalculateContinuumAbsorption 
         Caption         =   "Calculate Continuum Absorption"
         HelpContextID   =   226
      End
      Begin VB.Menu menuOutputCalculateChargeBalance 
         Caption         =   "Calculate Charge Balance"
         HelpContextID   =   227
      End
      Begin VB.Menu menuOutputCalculateTotalCations 
         Caption         =   "Calculate Total Cations"
         HelpContextID   =   686
      End
      Begin VB.Menu menuOutputSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu menuOutputDisplayAmphibole 
         Caption         =   "Display Amphibole Calculations"
         HelpContextID   =   631
      End
      Begin VB.Menu menuOutputDisplayBiotite 
         Caption         =   "Display Biotite Calculations"
         HelpContextID   =   632
      End
      Begin VB.Menu menuOutputSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu menuOutputDisplayZAFCalculation 
         Caption         =   "Display ZAF Calculations"
         HelpContextID   =   611
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "&Help"
      HelpContextID   =   228
      Begin VB.Menu menuHelpAboutStandard 
         Caption         =   "About STANDARD"
         HelpContextID   =   229
      End
      Begin VB.Menu menuHelpOnStandard 
         Caption         =   "Help on STANDARD"
         HelpContextID   =   230
         Shortcut        =   ^H
      End
      Begin VB.Menu menuHelpSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu menuHelpProbeSoftwareOnTheWeb 
         Caption         =   "Probe Software On The Web"
      End
      Begin VB.Menu menuHelpProbeSoftwareUserForum 
         Caption         =   "Connect To Probe Software User Forum"
      End
   End
End
Attribute VB_Name = "FormMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2016 by John J. Donovan
Option Explicit

Private Sub Form_Activate()
If Not DebugMode Then On Error Resume Next

Static initialized As Boolean

' Initialize the global variables
If initialized = False Then

' Initialize the files
Call InitFiles
If ierror Then End

' Load the PROBEWIN.INI file
Call InitINI
If ierror Then End

' Initialize arrays
Call InitData
If ierror Then End

' Init ZAF arrays
Call ZAFInitZAF
If ierror Then Exit Sub

' Initialize the log window
FormMAIN.TimerLogWindow.Interval = Int(LogWindowInterval! * 1000)

' Check for command line arguments and open the file if found (otherwise prompt user)
Call StanFormCommandLine(FormMAIN)
If ierror Then Exit Sub

' Check if the standard database needs to be updated
If StandardDataFile$ <> vbNullString Then
Call StandardUpdateMDBFile(StandardDataFile$)
If ierror Then Exit Sub
End If

' Update STANDARD FormMAIN
Call StanFormUpdate
If ierror Then Exit Sub

' Set default match to current database
If StandardDataFile$ <> vbNullString Then DefaultMatchStandardDatabase$ = StandardDataFile$
initialized = True
End If

End Sub

Private Sub Form_Load()
' Load FormMAIN for program STANDARD
If Not DebugMode Then On Error Resume Next

' Load form and application icon
Call MiscLoadIcon(FormMAIN)

' Check if program is already running
If app.PrevInstance Then
msg$ = "STANDARD is already running, click OK, then type <ctrl> <esc> for the Task Manager and select STANDARD.EXE from the Task List"
MsgBox msg$, vbOKOnly + vbExclamation, "STANDARD"
End
End If

Call InitWindow(Int(2), MDBUserName$, Me)

' Help file
FormMAIN.HelpContextID = IOGetHelpContextID("FormMAIN")

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
' Disable timer
FormMAIN.TimerLogWindow.Enabled = False

' Check for pending transactions
Call TransactionUnload("FormMAIN.Form_QueryUnload")
If ierror Then Exit Sub

End Sub

Private Sub Form_Resize()
If Not DebugMode Then On Error Resume Next
Dim tempsize As Integer

' Make text box (Log Window) full size of window
FormMAIN.TextLog.Left = 0
If FormMAIN.ScaleWidth > 0 Then FormMAIN.TextLog.Width = FormMAIN.ScaleWidth

tempsize% = FormMAIN.ScaleHeight - FormMAIN.TextLog.Top - FormMAIN.StatusBarAuto.Height
If tempsize% > 0 Then FormMAIN.TextLog.Height = tempsize%

' Move label controls
FormMAIN.Frame1.Width = FormMAIN.ScaleWidth - (FormMAIN.Frame2.Width + FRAMEBORDERWIDTH% * 3)
FormMAIN.Frame2.Left = FormMAIN.Frame1.Left + FormMAIN.Frame1.Width + FRAMEBORDERWIDTH%
FormMAIN.ListAvailableStandards.Width = FormMAIN.Frame1.Width - FRAMEBORDERWIDTH% * 2

End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End
End Sub

Private Sub ListAvailableStandards_DblClick()
If Not DebugMode Then On Error Resume Next
Dim stdnum As Integer

' Get standard from listbox
If FormMAIN.ListAvailableStandards.ListIndex < 0 Then Exit Sub
stdnum% = FormMAIN.ListAvailableStandards.ItemData(FormMAIN.ListAvailableStandards.ListIndex)

' Recalculate and display standard data
If stdnum% > 0 Then Call StanFormCalculate(stdnum%, Int(0))
If ierror Then Exit Sub

End Sub

Private Sub menuAnalyticalConditions_Click()
If Not DebugMode Then On Error Resume Next

' Load the form
Call CondLoad
If ierror Then Exit Sub

' Load COND form
FormCOND.Show vbModal

' Update MQOptions window if loaded
If FormMQOPTIONS.Visible = True Then
Call MqOptionsSave
If ierror Then Exit Sub
Call MqOptionsLoad
If ierror Then Exit Sub
End If

End Sub

Private Sub menuAnalyticalEmpiricalMACs_Click()
If Not DebugMode Then On Error Resume Next

' Load the empirical MAC/APFs
EmpTypeFlag% = 1

' Load form parameters
Call EmpLoad
If ierror Then Exit Sub

FormEMP.Show vbModal
If ierror Then Exit Sub

End Sub

Private Sub menuAnalyticalMQOptions_Click()
If Not DebugMode Then On Error Resume Next
Call MqOptionsLoad
If ierror Then Exit Sub
FormMQOPTIONS.Show vbModeless
If ierror Then Exit Sub
End Sub

Private Sub menuAnalyticalPENEPMA_Click()
If Not DebugMode Then On Error Resume Next
If Penepma08CheckPenepmaVersion%() = 6 Then
msg$ = "Penepma 2006 is no longer supported. Please download the latest PENEPMA12.ZIP or PENEPMA14.ZIP file and extract the files to the " & UserDataDirectory$ & " folder and check that the PENEPMA_Path, PENDBASE_Path and PENEPMA_Root strings are properly specified in the " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "menuAnayticalPENEPMA"
ElseIf Penepma08CheckPenepmaVersion%() = 8 Or Penepma08CheckPenepmaVersion%() = 12 Or Penepma08CheckPenepmaVersion%() = 14 Then
Call Penepma08Load(FormPENEPMA08_PE)
If ierror Then Exit Sub
Else
msg$ = "Penepma 2012 or 2014 application files were not found. Please download the PENEPMA12.ZIP or PENEPMA14.ZIP file and extract the files to the " & UserDataDirectory$ & " folder and check that the PENEPMA_Path, PENDBASE_Path and PENEPMA_Root strings are properly specified in the " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "menuAnayticalPENEPMA"
End If
End Sub

Private Sub menuAnalyticalPENFLUOR_Click()
If Penepma08CheckPenepmaVersion%() = 12 Or Penepma08CheckPenepmaVersion%() = 14 Then
Call Penepma12Load
If ierror Then Exit Sub
FormPENEPMA12.Show vbModeless
If ierror Then Exit Sub
Else
msg$ = "Penepma 2012 or 2014 application files were not found. Please download the PENEPMA12.ZIP or PENEPMA14.ZIP file and extract the files to the " & UserDataDirectory$ & " folder and check that the PENEPMA_Path, PENDBASE_Path and PENEPMA_Root strings are properly specified in the " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "menuAnayticalPENFLUOR"
End If
End Sub

Private Sub menuAnalyticalZAFSelections_Click()
If Not DebugMode Then On Error Resume Next

' Update or display ZAF selections
Call GetZAFAllLoad
If ierror Then Exit Sub
FormGETZAFALL.Option6(4).Enabled = False    ' disable calibration curve
FormGETZAFALL.Show vbModal
If ierror Then Exit Sub

Call TypeZAFSelections
If ierror Then Exit Sub
End Sub

Private Sub menuEditClearAll_Click()
If Not DebugMode Then On Error Resume Next
FormMAIN.TextLog.Text = vbNullString
End Sub

Private Sub menuEditCopy_Click()
If Not DebugMode Then On Error Resume Next
Clipboard.Clear
Clipboard.SetText FormMAIN.TextLog.SelText
End Sub

Private Sub menuEditCut_Click()
If Not DebugMode Then On Error Resume Next
Clipboard.Clear
Clipboard.SetText FormMAIN.TextLog.SelText
FormMAIN.TextLog.SelText = vbNullString
End Sub

Private Sub menuEditPaste_Click()
If Not DebugMode Then On Error Resume Next
FormMAIN.TextLog.SelText = Clipboard.GetText()
End Sub

Private Sub menuEditSelectAll_Click()
If Not DebugMode Then On Error Resume Next
FormMAIN.TextLog.SetFocus
FormMAIN.TextLog.SelStart = 0
FormMAIN.TextLog.SelLength = Len(FormMAIN.TextLog.Text)
End Sub

Private Sub menuFileAMCSD_Click()
If Not DebugMode Then On Error Resume Next
Dim response As Integer
Dim tfilename As String

If StandardDataFile$ <> vbNullString Then
msg$ = "Do you want to close the currently open standard database and create a new AMCSD.MDB mineral database?"
response% = MsgBox(msg$, vbYesNo + vbQuestion + vbDefaultButton1, "Standard")
If response% = vbNo Then Exit Sub
End If

' Read the PROBEWIN.INI file
Call InitINI

' Initialize the arrays
Call InitData
Call InitStandardIndex

' Clear the form
Call StanFormClear
Call StanFormUpdate

' Open new file for importing American Mineralogist data
tfilename$ = "amcsd.mdb"
Call StandardOpenNEWFile(tfilename$, FormMAIN)
If ierror Then Exit Sub

' Update the file info fields
MDBUserName$ = "John Donovan"
MDBFileTitle$ = "American Mineralogist Mineral Formula Database"
MDBFileDescription$ = "From Andy McCarthy, University of Arizona" & vbCrLf & "Project RRUFF, 2006"
Call FileInfoSaveData(tfilename$)
If ierror Then Exit Sub

' Open the import file
ImportDataFile$ = ApplicationCommonAppData$ & "all_weights-modified.txt"
Open ImportDataFile$ For Input As #ImportDataFileNumber%

' Import the standards (close file before exiting on error)
Call IOStatusAuto(vbNullString)
Call StandardReadDATFile(Int(3))
Call IOStatusAuto(vbNullString)
Close #ImportDataFileNumber%    ' close import file

' Load the standard list box (even if there was an error importing)
Call StandardLoadList(FormMAIN.ListAvailableStandards)
If ierror Then Exit Sub

' Update the MAIN form
Call StanFormUpdate
If ierror Then Exit Sub

End Sub

Private Sub menuFileClose_Click()
If Not DebugMode Then On Error Resume Next

' Close .OUT file if open
Call IOCloseOUTFile

' Read the PROBEWIN.INI file
Call InitINI

' Initialize the arrays
Call InitData
Call InitStandardIndex

' Blank standard file
StandardDataFile$ = vbNullString
FormMAIN.menuFileClose.Enabled = False
FormMAIN.menuFileNew.Enabled = True
FormMAIN.menuFileOpen.Enabled = True

' Clear the form
Call StanFormClear
Call StanFormUpdate

' Close some forms
Unload FormPENEPMA08_PE
Unload FormPENEPMA08Batch
Unload FormPENEPMA12
Unload FormPenepma12Binary
Unload FormPenepma12Random

Unload FormMATCH
Unload FormMODAL
Unload FormFIND
Unload FormFIND2
Unload FormMQOPTIONS
Unload FormXRAY

End Sub

Private Sub menuFileExit_Click()
If Not DebugMode Then On Error Resume Next
Unload FormMAIN
End Sub

Private Sub menuFileExport_Click()
If Not DebugMode Then On Error Resume Next

' Open the standard export file
Call StandardOpenDATFile(Int(1), FormMAIN.ListAvailableStandards, FormMAIN)
If ierror Then Exit Sub

' Export the standards
Call IOStatusAuto(vbNullString)
Call StandardWriteDATFile(Int(1))
Call IOStatusAuto(vbNullString)
Close #ImportDataFileNumber%

End Sub

Private Sub menuFileExportSingleRowFormat_Click()
If Not DebugMode Then On Error Resume Next

' Open the standard export file
Call StandardOpenDATFile(Int(3), FormMAIN.ListAvailableStandards, FormMAIN)
If ierror Then Exit Sub

' Export the standards
Call IOStatusAuto(vbNullString)
Call StandardWriteDATFile(Int(2))
Call IOStatusAuto(vbNullString)
Close #ImportDataFileNumber%

' Ask if export to Excel
Call StandardWriteDATFile2(FormMAIN)
If ierror Then Exit Sub

End Sub

Private Sub menuFileFileInformation_Click()
If Not DebugMode Then On Error Resume Next

' Load database file info
Call FileInfoLoad(Int(1), StandardDataFile$)
If ierror Then Exit Sub

FormFILEINFO.Show vbModal
If ierror Then Exit Sub

Call StanFormUpdate
If ierror Then Exit Sub

End Sub

Private Sub menuFileImport_Click()
If Not DebugMode Then On Error Resume Next

' Open the standard import file
Call StandardOpenDATFile(Int(2), FormMAIN.ListAvailableStandards, FormMAIN)
If ierror Then Exit Sub

' Import the standards (close file before exiting on error)
Call IOStatusAuto(vbNullString)
Call StandardReadDATFile(Int(1))
Call IOStatusAuto(vbNullString)
Close #ImportDataFileNumber%    ' close import file

' Load the standard list box (even if there was an error importing)
Call StandardLoadList(FormMAIN.ListAvailableStandards)
If ierror Then Exit Sub

' Update the MAIN form
Call StanFormUpdate
If ierror Then Exit Sub

End Sub

Private Sub menuFileImportSingleRowFormat_Click()
If Not DebugMode Then On Error Resume Next

' Open the standard import file
Call StandardOpenDATFile(Int(4), FormMAIN.ListAvailableStandards, FormMAIN)
If ierror Then Exit Sub

' Import the standards (close file before exiting on error)
Call IOStatusAuto(vbNullString)
Call StandardReadDATFile(Int(2))
Call IOStatusAuto(vbNullString)
Close #ImportDataFileNumber%    ' close import file

' Load the standard list box (even if there was an error importing)
Call StandardLoadList(FormMAIN.ListAvailableStandards)
If ierror Then Exit Sub

' Update the MAIN form
Call StanFormUpdate
If ierror Then Exit Sub

End Sub

Private Sub menuFileImportStandardsFromCamecaPeakSight_Click()
If Not DebugMode Then On Error Resume Next
Dim response As Integer
Dim tfilename As String

If StandardDataFile$ <> vbNullString Then
msg$ = "Do you want to import standard compositions from the Cameca SX.MDB file, into the currently open standard database?"
response% = MsgBox(msg$, vbYesNo + vbQuestion + vbDefaultButton2, "Standard")
If response% = vbNo Then Exit Sub
End If

' Read the PROBEWIN.INI file
Call InitINI

' Initialize the arrays
Call InitData
Call InitStandardIndex

' Clear the form
Call StanFormClear
Call StanFormUpdate

' Open new file for importing Cameca PeakSight data if none currently open
If StandardDataFile$ = vbNullString Then
tfilename$ = "standard.mdb"
Call StandardOpenNEWFile(tfilename$, FormMAIN)
If ierror Then Exit Sub

' Update the file info fields
MDBUserName$ = MDBUserName$
MDBFileTitle$ = "Default Standard Database"
MDBFileDescription$ = "Cameca PeakSight Import of Sx.mdb database"
Call FileInfoSaveData(tfilename$)
If ierror Then Exit Sub
End If

' Import the Cameca standards
Call StandardImportCameca(FormMAIN)
'If ierror Then Exit Sub

' Load the standard list box (even if there was an error importing)
Call StandardLoadList(FormMAIN.ListAvailableStandards)
If ierror Then Exit Sub

' Update the MAIN form
Call StanFormUpdate
If ierror Then Exit Sub

End Sub

Private Sub menuFileImportStandardsFromJEOLTextFile_Click()
If Not DebugMode Then On Error Resume Next
Dim response As Integer
Dim tfilename As String

If StandardDataFile$ <> vbNullString Then
msg$ = "Do you want to close the currently open standard database and create a new STANDARD.MDB default database from the JEOL Import Text File?"
response% = MsgBox(msg$, vbYesNo + vbQuestion + vbDefaultButton2, "Standard")
If response% = vbNo Then Exit Sub
End If

' Read the PROBEWIN.INI file
Call InitINI

' Initialize the arrays
Call InitData
Call InitStandardIndex

' Clear the form
Call StanFormClear
Call StanFormUpdate

' Open new file for importing JEOL standard data
tfilename$ = "standard.mdb"
Call StandardOpenNEWFile(tfilename$, FormMAIN)
If ierror Then Exit Sub

' Update the file info fields
MDBUserName$ = MDBUserName$
MDBFileTitle$ = "Default Standard Database"
MDBFileDescription$ = "JEOL Text File Import Standard Database"
Call FileInfoSaveData(tfilename$)
If ierror Then Exit Sub

' Import the JEOL standards
Call StandardImportJEOL(FormMAIN)
'If ierror Then Exit Sub

' Load the standard list box (even if there was an error importing)
Call StandardLoadList(FormMAIN.ListAvailableStandards)
If ierror Then Exit Sub

' Update the MAIN form
Call StanFormUpdate
If ierror Then Exit Sub

End Sub

Private Sub menuFileInputCalcZAFStandardFormatKratios_Click()
If Not DebugMode Then On Error Resume Next
Call StandardInputCalcZAFStandardFormatKratios(FormMAIN)
If ierror Then Exit Sub
End Sub

Private Sub menuFileNew_Click()
If Not DebugMode Then On Error Resume Next

' Open a new standard file
Call StandardOpenNEWFile(vbNullString, FormMAIN)
If ierror Then Exit Sub

' Confirm file info
Call FileInfoLoad(Int(1), StandardDataFile$)
If ierror Then Exit Sub
FormFILEINFO.Show vbModal

Call StanFormClear
Call StanFormUpdate

FormMAIN.menuFileClose.Enabled = True
FormMAIN.menuFileNew.Enabled = False
FormMAIN.menuFileOpen.Enabled = False

End Sub

Private Sub menuFileOpen_Click()
If Not DebugMode Then On Error Resume Next
Dim tfilename As String

' Open an existing file (don't exit on error)
Call StandardOpenMDBFile(tfilename$, FormMAIN)
If ierror Then Exit Sub

' Update the form
Call StanFormUpdate
If ierror Then Exit Sub

FormMAIN.menuFileClose.Enabled = True
FormMAIN.menuFileNew.Enabled = False
FormMAIN.menuFileOpen.Enabled = False
End Sub

Private Sub menuFilePrintLog_Click()
If Not DebugMode Then On Error Resume Next
' Print dialog
Call IOPrintLog
If ierror Then Exit Sub
End Sub

Private Sub menuFilePrintSetup_Click()
If Not DebugMode Then On Error Resume Next
' Print setup dialog
Call IOPrintSetup
If ierror Then Exit Sub
End Sub

Private Sub menuFileSaveAs_Click()
If Not DebugMode Then On Error Resume Next
' Save standard database to a new filename
Call StanFormSaveAsFile(FormMAIN)
If ierror Then Exit Sub
End Sub

Private Sub menuHelpAboutStandard_Click()
If Not DebugMode Then On Error Resume Next
FormABOUT.Show vbModal
End Sub

Private Sub menuHelpOnStandard_Click()
If Not DebugMode Then On Error Resume Next
Call MiscFormLoadHelp(FormMAIN.HelpContextID)
If ierror Then Exit Sub
End Sub

Private Sub menuHelpProbeSoftwareOnTheWeb_Click()
If Not DebugMode Then On Error Resume Next
Call IOBrowseHTTP(ProbeSoftwareInternetBrowseMethod%, "http://probesoftware.com/index.html")
If ierror Then Exit Sub
End Sub

Private Sub menuHelpProbeSoftwareUserForum_Click()
If Not DebugMode Then On Error Resume Next
Call IOBrowseHTTP(ProbeSoftwareInternetBrowseMethod%, "http://probesoftware.com/smf/index.php")
If ierror Then Exit Sub
End Sub

Private Sub menuOptionsFind_Click()
If Not DebugMode Then On Error Resume Next
' Load the find standard dialog
Call FindLoad
If ierror Then Exit Sub
FormFIND.Show vbModeless
End Sub

Private Sub menuOptionsInterferences_Click()
If Not DebugMode Then On Error Resume Next
' Load the nominal interference calculation dialog
Call InterfLoad
If ierror Then Exit Sub
FormINTERF.Show vbModeless
End Sub

Private Sub menuOptionsMatch_Click()
If Not DebugMode Then On Error Resume Next
' Load the match composition dialog
Call StandardMatchLoad
If ierror Then Exit Sub
FormMATCH.Show vbModeless
End Sub

Private Sub menuOptionsModalAnalysis_Click()
If Not DebugMode Then On Error Resume Next
Call ModalLoadForm
If ierror Then Exit Sub
FormMODAL.Show vbModeless
End Sub

Private Sub menuOptionsSearch_Click()
If Not DebugMode Then On Error Resume Next
FormFIND2.Show vbModeless
End Sub

Private Sub menuOutputCalculateAlternativeZbars_Click()
If Not DebugMode Then On Error Resume Next
FormMAIN.menuOutputCalculateAlternativeZbars.Checked = Not FormMAIN.menuOutputCalculateAlternativeZbars.Checked
CalculateAlternativeZbarsFlag = FormMAIN.menuOutputCalculateAlternativeZbars.Checked
End Sub

Private Sub menuOutputCalculateChargeBalance_Click()
If Not DebugMode Then On Error Resume Next
FormMAIN.menuOutputCalculateChargeBalance.Checked = Not FormMAIN.menuOutputCalculateChargeBalance.Checked
UseChargeBalanceCalculationFlag = FormMAIN.menuOutputCalculateChargeBalance.Checked
End Sub

Private Sub menuOutputCalculateContinuumAbsorption_Click()
If Not DebugMode Then On Error Resume Next
FormMAIN.menuOutputCalculateContinuumAbsorption.Checked = Not FormMAIN.menuOutputCalculateContinuumAbsorption.Checked
CalculateContinuumAbsorptionFlag = FormMAIN.menuOutputCalculateContinuumAbsorption.Checked
End Sub

Private Sub menuOutputCalculateElectronandXrayRanges_Click()
If Not DebugMode Then On Error Resume Next
FormMAIN.menuOutputCalculateElectronandXrayRanges.Checked = Not FormMAIN.menuOutputCalculateElectronandXrayRanges.Checked
CalculateElectronandXrayRangesFlag = FormMAIN.menuOutputCalculateElectronandXrayRanges.Checked
End Sub

Private Sub menuOutputCalculateTotalCations_Click()
If Not DebugMode Then On Error Resume Next
FormMAIN.menuOutputCalculateTotalCations.Checked = Not FormMAIN.menuOutputCalculateTotalCations.Checked
UseTotalCationsCalculationFlag = FormMAIN.menuOutputCalculateTotalCations.Checked
End Sub

Private Sub menuOutputDebugMode_Click()
If Not DebugMode Then On Error Resume Next
FormMAIN.menuOutputDebugMode.Checked = Not FormMAIN.menuOutputDebugMode.Checked
DebugMode = FormMAIN.menuOutputDebugMode.Checked
End Sub

Private Sub menuOutputDisplayAmphibole_Click()
If Not DebugMode Then On Error Resume Next
FormMAIN.menuOutputDisplayAmphibole.Checked = Not FormMAIN.menuOutputDisplayAmphibole.Checked
DisplayAmphiboleCalculationFlag = FormMAIN.menuOutputDisplayAmphibole.Checked
End Sub

Private Sub menuOutputDisplayBiotite_Click()
If Not DebugMode Then On Error Resume Next
FormMAIN.menuOutputDisplayBiotite.Checked = Not FormMAIN.menuOutputDisplayBiotite.Checked
DisplayBiotiteCalculationFlag = FormMAIN.menuOutputDisplayBiotite.Checked
End Sub

Private Sub menuOutputDisplayZAFCalculation_Click()
If Not DebugMode Then On Error Resume Next
FormMAIN.menuOutputDisplayZAFCalculation.Checked = Not FormMAIN.menuOutputDisplayZAFCalculation.Checked
DisplayZAFCalculationFlag = FormMAIN.menuOutputDisplayZAFCalculation.Checked
End Sub

Private Sub menuOutputExtendedFormat_Click()
If Not DebugMode Then On Error Resume Next
FormMAIN.menuOutputExtendedFormat.Checked = Not FormMAIN.menuOutputExtendedFormat.Checked
ExtendedFormat = FormMAIN.menuOutputExtendedFormat.Checked
End Sub

Private Sub menuOutputLogWindow_Click()
If Not DebugMode Then On Error Resume Next
' Load the change log window font dialog
Call IOLogFont
If ierror Then Exit Sub
End Sub

Private Sub menuOutputSaveToDiskLog_Click()
' This routine toggles the "SaveToDisk" flag and opens or closes the .OUT file
If Not DebugMode Then On Error Resume Next

' Perform file operations and update flag
If Not SaveToDisk Then
Call IOOpenOUTFile(FormMAIN)
Else
Call IOCloseOUTFile
End If

End Sub

Private Sub menuOutputVerboseMode_Click()
If Not DebugMode Then On Error Resume Next
FormMAIN.menuOutputVerboseMode.Checked = Not FormMAIN.menuOutputVerboseMode.Checked
VerboseMode = FormMAIN.menuOutputVerboseMode.Checked
End Sub

Private Sub menuStandardDelete_Click()
If Not DebugMode Then On Error Resume Next
' Delete a single standard
Call StanFormDeleteSingleStandard
If ierror Then Exit Sub
End Sub

Private Sub menuStandardDeleteSelected_Click()
If Not DebugMode Then On Error Resume Next
' Delete selected standards
Call StanFormDeleteSelectedStandards
If ierror Then Exit Sub
End Sub

Private Sub menuStandardDuplicate_Click()
If Not DebugMode Then On Error Resume Next
' Duplicate the currently selected standard
Call StanFormDuplicate
If ierror Then Exit Sub
End Sub

Private Sub menuStandardListAllStandards_Click()
If Not DebugMode Then On Error Resume Next
Call IOStatusAuto(vbNullString)
Call StanFormListStandards(Int(2))
Call IOStatusAuto(vbNullString)
If ierror Then Exit Sub
End Sub

Private Sub menuStandardListElementalStandardNames_Click()
If Not DebugMode Then On Error Resume Next
Call StanFormListNames(Int(1), Int(0))
If ierror Then Exit Sub
End Sub

Private Sub menuStandardListElementalStandards_Click()
If Not DebugMode Then On Error Resume Next
Call IOStatusAuto(vbNullString)
Call StanFormListStandards(Int(3))
Call IOStatusAuto(vbNullString)
If ierror Then Exit Sub
End Sub

Private Sub menuStandardListOxideStandardNames_Click()
If Not DebugMode Then On Error Resume Next
Call StanFormListNames(Int(2), Int(0))
If ierror Then Exit Sub
End Sub

Private Sub menuStandardListOxideStandards_Click()
If Not DebugMode Then On Error Resume Next
Call IOStatusAuto(vbNullString)
Call StanFormListStandards(Int(4))
Call IOStatusAuto(vbNullString)
If ierror Then Exit Sub
End Sub

Private Sub menuStandardListSelectedStandards_Click()
If Not DebugMode Then On Error Resume Next
Call IOStatusAuto(vbNullString)
Call StanFormListStandards(Int(1))
Call IOStatusAuto(vbNullString)
If ierror Then Exit Sub
End Sub

Private Sub menuStandardListStandardNames_Click()
If Not DebugMode Then On Error Resume Next
' List all standard names in database
Call StanFormListNames(Int(0), Int(0))
If ierror Then Exit Sub
End Sub

Private Sub menuStandardListStandardNamesZbar_Click()
If Not DebugMode Then On Error Resume Next
' List all standard names in database with average Z
Call StanFormListNames(Int(0), Int(1))
If ierror Then Exit Sub
End Sub

Private Sub menuStandardModify_Click()
If Not DebugMode Then On Error Resume Next
' Modify the currently selected standard
Call StanFormModify
If ierror Then Exit Sub
End Sub

Private Sub menuStandardNew_Click()
If Not DebugMode Then On Error Resume Next
' Create a new standard in the current standard database
Call StanFormNew
If ierror Then Exit Sub
End Sub

Private Sub menuViewDiskLog_Click()
If Not DebugMode Then On Error Resume Next
' View disk log file
Call IOViewLog
If ierror Then Exit Sub
End Sub

Private Sub menuXrayEdgeTable_Click()
If Not DebugMode Then On Error Resume Next
' Obtain an edge table
Call XrayGetTable(Int(2))
If ierror Then Exit Sub
End Sub

Private Sub menuXrayEmissionTable_Click()
If Not DebugMode Then On Error Resume Next
' Obtain an emission table
Call XrayGetTable(Int(1))
If ierror Then Exit Sub
End Sub

Private Sub menuXrayFluorescentYieldTable_Click()
If Not DebugMode Then On Error Resume Next
' Obtain an fluorescent yield table
Call XrayGetTable(Int(3))
If ierror Then Exit Sub
End Sub

Private Sub menuXrayMACTable_Click()
If Not DebugMode Then On Error Resume Next
' Obtain a MAC table
Call XrayGetTable(Int(7))
If ierror Then Exit Sub
End Sub

Private Sub menuXrayXrayDatabase_Click()
If Not DebugMode Then On Error Resume Next
' Loads the xray database
Call XrayGetDatabase
If ierror Then Exit Sub
End Sub

Private Sub StatusBarAuto_PanelClick(ByVal Panel As ComctlLib.Panel)
If Not DebugMode Then On Error Resume Next
Select Case Panel.Key
Case "status"
    Exit Sub
Case "cancel"
    Call IOAutomationCancel
    If ierror Then Exit Sub
Case "pause"
    Call IOAutomationPause(Int(0))
    If ierror Then Exit Sub
Case Else
    Exit Sub
End Select
End Sub

Private Sub TextLog_KeyPress(KeyAscii As Integer)
If Not DebugMode Then On Error Resume Next
Call IOSendLog(KeyAscii%)
If ierror Then Exit Sub
End Sub

Private Sub TimerLogWindow_Timer()
If Not DebugMode Then On Error Resume Next
Call IODumpLog
End Sub

