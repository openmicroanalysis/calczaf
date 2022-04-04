VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Object = "{7ED47906-67D7-4D60-ABCD-66C3BA9E3452}#1.0#0"; "csmtpctl.ocx"
Object = "{959AC9FE-B2CE-4117-9CE6-56B273C5848F}#1.0#0"; "csmsgctl.ocx"
Begin VB.Form FormMAIN 
   Caption         =   "CalcZAF (Calculate ZAF and Phi-Rho-Z Corrections)"
   ClientHeight    =   4920
   ClientLeft      =   690
   ClientTop       =   3060
   ClientWidth     =   10605
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
   Icon            =   "CALCZAF.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4920
   ScaleWidth      =   10605
   Begin MailMessageCtl.MailMessage MailMessage1 
      Left            =   2280
      Top             =   0
      _cx             =   741
      _cy             =   741
   End
   Begin SmtpClientCtl.SmtpClient SmtpClient1 
      Left            =   1680
      Top             =   0
      _cx             =   741
      _cy             =   741
   End
   Begin ComctlLib.StatusBar StatusBarAuto 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4665
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   14579
            TextSave        =   ""
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
   Begin RichTextLib.RichTextBox TextLog 
      Height          =   4335
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   7646
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"CALCZAF.frx":59D8A
   End
   Begin VB.Label LabelTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label LabelCalculated 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label LabelAtomic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label LabelTotalOxygen 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7320
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label LabelExcess 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7320
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label LabelZbar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7320
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Menu menuFile 
      Caption         =   "&File"
      HelpContextID   =   98
      Begin VB.Menu menuFileOpen 
         Caption         =   "&Open CalcZAF Input Data File (Test with CALCZAF.DAT, CALCBIN.DAT, AuCu_NBS-K-ratios.DAT, Olivine particle-JTA-0.5um.DAT, etc.)"
         HelpContextID   =   715
      End
      Begin VB.Menu menuFileClose 
         Caption         =   "&Close CalcZAF Input Data File"
         HelpContextID   =   716
      End
      Begin VB.Menu menuFileSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu menuFileOpenAndProcess 
         Caption         =   "Open Input Data File And Calculate/Export All"
      End
      Begin VB.Menu menuFileSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu menuFileExport 
         Caption         =   "Export CalcZAF Input Data File"
         HelpContextID   =   598
      End
      Begin VB.Menu menuFileUpdateCalcZAFSampleDataFiles 
         Caption         =   "Update CalcZAF Example Data Files"
         HelpContextID   =   768
      End
      Begin VB.Menu menuFileSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu menuFilePrintLog 
         Caption         =   "&Print Log"
         HelpContextID   =   101
      End
      Begin VB.Menu menuFilePrintSetup 
         Caption         =   "Print Setup"
         HelpContextID   =   102
      End
      Begin VB.Menu menuFileSeparator4 
         Caption         =   "-"
      End
      Begin VB.Menu menuFileExit 
         Caption         =   "Exit"
         HelpContextID   =   103
      End
   End
   Begin VB.Menu menuEdit 
      Caption         =   "&Edit"
      HelpContextID   =   104
      Begin VB.Menu menuEditCut 
         Caption         =   "Cut"
         HelpContextID   =   105
      End
      Begin VB.Menu menuEditCopy 
         Caption         =   "Copy"
         HelpContextID   =   106
      End
      Begin VB.Menu menuEditPaste 
         Caption         =   "Paste"
         HelpContextID   =   107
      End
      Begin VB.Menu menuEditSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu menuEditSelectAll 
         Caption         =   "Select All"
         HelpContextID   =   108
      End
      Begin VB.Menu menuClearAll 
         Caption         =   "Clear All"
         HelpContextID   =   109
      End
   End
   Begin VB.Menu menuStandard 
      Caption         =   "&Standard"
      HelpContextID   =   110
      Begin VB.Menu menuStandardStandardDatabase 
         Caption         =   "&Standard Database"
         HelpContextID   =   111
      End
      Begin VB.Menu menuStandardSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu menuStandardSelectStandardDatabase 
         Caption         =   "Select Standard Database"
         HelpContextID   =   717
      End
      Begin VB.Menu menuStandardEditStandardParameters 
         Caption         =   "&Edit Standard Coating Parameters"
         HelpContextID   =   719
      End
      Begin VB.Menu menuStandardSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu menuStandardAddStandardsToRun 
         Caption         =   "&Add/Remove Standards To/From Run"
         HelpContextID   =   796
      End
   End
   Begin VB.Menu menuXray 
      Caption         =   "&X-Ray"
      HelpContextID   =   113
      Begin VB.Menu menuXrayXrayDatabase 
         Caption         =   "&X-Ray Database"
         HelpContextID   =   114
      End
      Begin VB.Menu menuXrayCalculateSpectrometerPosition 
         Caption         =   "&Calculate Spectrometer Position"
      End
      Begin VB.Menu menuXraySeparator0 
         Caption         =   "-"
      End
      Begin VB.Menu menuXrayEmissionTable 
         Caption         =   "Emission Table (Ka, Kb, La, Lb, Ma, Mb)"
         HelpContextID   =   115
      End
      Begin VB.Menu menuXrayEdgeTable 
         Caption         =   "Edge Table (K, L-I, L-II, L-III, M-I, M-II, M-III, M-IV, M-V)"
         HelpContextID   =   116
      End
      Begin VB.Menu menuXrayFluorescentYieldTable 
         Caption         =   "Fluorescent Yield Table (Ka, Kb, La, Lb, Ma, Mb)"
         HelpContextID   =   117
      End
      Begin VB.Menu menuXraySeparator00 
         Caption         =   "-"
      End
      Begin VB.Menu menuXrayEmissionTable2 
         Caption         =   "Emission Table (Ln, Lg, Lv, Ll, Mg, Mz)"
      End
      Begin VB.Menu menuXrayFluorescentYieldTable2 
         Caption         =   "Fluorescent Yield Table (Ln, Lg, Lv, Ll, Mg, Mz)"
      End
      Begin VB.Menu menuXraySeparator000 
         Caption         =   "-"
      End
      Begin VB.Menu menuXrayMACTable 
         Caption         =   "MAC Table (Ka, Kb, La, Lb, Ma, Mb)"
         HelpContextID   =   118
      End
      Begin VB.Menu menuXrayMACTableComplete 
         Caption         =   "MAC Table (complete) (Ka, Kb, La, Lb, Ma, Mb)"
         HelpContextID   =   509
      End
      Begin VB.Menu menuXraySeparator111 
         Caption         =   "-"
      End
      Begin VB.Menu menuXrayMACTable2 
         Caption         =   "MAC Table (Ln, Lg, Lv, Ll, Mg, Mz)"
      End
      Begin VB.Menu menuXrayMACTableComplete2 
         Caption         =   "MAC Table (complete) (Ln, Lg, Lv, Ll, Mg, Mz)"
      End
      Begin VB.Menu menuXraySeparator1111 
         Caption         =   "-"
      End
      Begin VB.Menu menuXrayDisplayMACEmitterAbsorber 
         Caption         =   "&Display MAC Emitter Absorber Pair"
         HelpContextID   =   510
      End
      Begin VB.Menu menuXraySeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu menuXrayEditXrayTable 
         Caption         =   "Edit XLINE.DAT Table"
         HelpContextID   =   119
      End
      Begin VB.Menu menuXrayEditXedgeTable 
         Caption         =   "Edit XEDGE.DAT Table"
         HelpContextID   =   120
      End
      Begin VB.Menu menuXrayEditXflurTable 
         Caption         =   "Edit XFLUR.DAT Table"
         HelpContextID   =   121
      End
      Begin VB.Menu menuXraySeparator10 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu menuXrayUpdateXLineTable 
         Caption         =   "Update XLINE.DAT Table (from XLINE.TXT)"
         Visible         =   0   'False
      End
      Begin VB.Menu menuXrayUpdateXEdgeTable 
         Caption         =   "Update XEDGE.DAT Table (from XEDGE.TXT)"
         Visible         =   0   'False
      End
      Begin VB.Menu menuXrayUpdateXFlurTable 
         Caption         =   "Update XFLUR.DAT Table (from XFLUR.TXT)"
         Visible         =   0   'False
      End
      Begin VB.Menu menuXraySeparator11 
         Caption         =   "-"
      End
      Begin VB.Menu menuXrayEditMACTable 
         Caption         =   "&Edit MAC Table"
         HelpContextID   =   122
      End
      Begin VB.Menu menuXraySeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu menuXrayConvertELEMINFODAT 
         Caption         =   "Convert JTA ELEMINFO.DAT (Create New XLINE.DAT, XEDGE.DAT and XFLUR.DAT)"
         HelpContextID   =   123
      End
      Begin VB.Menu menuXraySeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu menuXrayConvertMACMATDAT 
         Caption         =   "Create New CITZMU MAC Table (Create New CITZMU.DAT)"
         HelpContextID   =   124
      End
      Begin VB.Menu menuXrayCreateNewMcMasterMACTable 
         Caption         =   "Create New McMaster MAC Table (Create New MCMASTER.DAT)"
         HelpContextID   =   126
      End
      Begin VB.Menu menuXrayCreateNewMAC30MACTable 
         Caption         =   "Create New MAC30 MAC Table (Create New MAC30.DAT)"
         HelpContextID   =   127
      End
      Begin VB.Menu menuXrayCreateNewMACJTAMACTable 
         Caption         =   "Create New MACJTA MAC Table (Create New MACJTA.DAT)"
         HelpContextID   =   128
      End
      Begin VB.Menu menuXrayCreateNewFFASTMACTable 
         Caption         =   "Create New FFAST MAC Table (Create New FFAST.DAT)"
         HelpContextID   =   576
      End
      Begin VB.Menu menuXraySeparator33 
         Caption         =   "-"
      End
      Begin VB.Menu menuXrayCreateNewUSERMACTable 
         Caption         =   "Create Default User Defined MAC Table (Create Default USERMAC.DAT from another MAC file)"
         HelpContextID   =   825
      End
      Begin VB.Menu menuXrayOutputExistingUSERMACTable 
         Caption         =   "Output Exisiting User Defined MAC Table (Create New USERMAC.TXT and USERMAC2.TXT)"
      End
      Begin VB.Menu menuXrayUpdateUSERMACTable 
         Caption         =   "Update Existing User Defined MAC Table (Update USERMAC.DAT and USERMAC2.DAT)"
         HelpContextID   =   826
      End
      Begin VB.Menu menuXraySeparator7 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu menuXrayUpdateEdgeLineFlurFiles 
         Caption         =   "Update XLINE, XEDGE, XFLUR .DAT Files For Actinides"
         HelpContextID   =   628
         Visible         =   0   'False
      End
      Begin VB.Menu menuXrayConvertTextToData 
         Caption         =   "Convert XLINE, XEDGE, XFLUR .TXT Files To Binary .DAT Files"
         HelpContextID   =   630
         Visible         =   0   'False
      End
      Begin VB.Menu menuXrayConvertDataToText 
         Caption         =   "Convert Binary XLINE, XEDGE, XFLUR .DAT Files To .TXT Files"
         Visible         =   0   'False
      End
      Begin VB.Menu menuXraySeparator8 
         Caption         =   "-"
      End
      Begin VB.Menu menuXrayCreateNewXrayDatabase 
         Caption         =   "Create New X-Ray Database (XRAY.ALL -> XRAY.MDB)"
         HelpContextID   =   125
      End
   End
   Begin VB.Menu menuAnalytical 
      Caption         =   "&Analytical"
      HelpContextID   =   129
      Begin VB.Menu menuAnalyticalZAFSelections 
         Caption         =   "&ZAF, Phi-Rho-Z, Alpha Factor and Calibration Curve Selections"
         HelpContextID   =   134
      End
      Begin VB.Menu menuAnalyticalConditions 
         Caption         =   "&Operating Conditions"
         HelpContextID   =   135
      End
      Begin VB.Menu menuAnalyticalEmpiricalMACs 
         Caption         =   "&Empirical MACs"
         HelpContextID   =   133
      End
      Begin VB.Menu menuAnalyticalParticleandThinFilm 
         Caption         =   "&Particle and Thin Film"
         HelpContextID   =   528
      End
      Begin VB.Menu menuAnalyticalElements 
         Caption         =   "Display Current Sample Elements"
         HelpContextID   =   770
      End
      Begin VB.Menu menuAnalyticalSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu menuAnalyticalModelSeco 
         Caption         =   $"CALCZAF.frx":59E0E
      End
      Begin VB.Menu menuAnalyticalSecondary 
         Caption         =   "Correct Secondary Fluorescence Boundary Effects (correct intensities exported from PFE as ASCII file)"
      End
      Begin VB.Menu menuAnalyticalSeparator0 
         Caption         =   "-"
      End
      Begin VB.Menu menuAnalyticalElementalFactors 
         Caption         =   "Calculate Elemental To Oxide Factors"
         HelpContextID   =   131
      End
      Begin VB.Menu menuAnalyticalOxideFactors 
         Caption         =   "Calculate Oxide To Elemental Factors"
         HelpContextID   =   132
      End
      Begin VB.Menu menuStudentstTable 
         Caption         =   "Students ""t"" Table"
         HelpContextID   =   626
      End
      Begin VB.Menu menuAnalyticalSeparator11 
         Caption         =   "-"
      End
      Begin VB.Menu menuAnalyticalUseConductiveCoatingCorrectionForElectronAbsorption 
         Caption         =   "Use Conductive Coating Correction For Electron Absorption"
         HelpContextID   =   773
      End
      Begin VB.Menu menuAnalyticalUseConductiveCoatingCorrectionForXrayTransmission 
         Caption         =   "Use Conductive Coating Correction For Xray Transmission"
         HelpContextID   =   774
      End
      Begin VB.Menu menuAnalyticalSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu menuAnalyticalAlphaFactors 
         Caption         =   "&Calculate and Plot Binary Alpha Factors"
         HelpContextID   =   627
      End
      Begin VB.Menu menuAnalyticalKFactorsAlphaFactors 
         Caption         =   "Calculate and Ouput Binary K-Ratios and Alpha Factors for Periodic Table"
         HelpContextID   =   653
      End
      Begin VB.Menu menuAnalyticalSeparator4 
         Caption         =   "-"
      End
      Begin VB.Menu menuAnalyticalBinaryCalculationOptions 
         Caption         =   "&Binary Calculation Options"
         HelpContextID   =   136
      End
      Begin VB.Menu menuAnalyticalCalculateBinaryIntensities 
         Caption         =   $"CALCZAF.frx":59EA7
         HelpContextID   =   511
      End
      Begin VB.Menu menuAnalyticalCalculateBinaryIntensitiesAllCorrections 
         Caption         =   "Calculate Binary Intensities (Output Calculated Intensities for All Matrix Corrections and MAC Files, Single Line, Single File)"
         HelpContextID   =   139
      End
      Begin VB.Menu menuAnalyticalCalculateBinaryIntensitiesAllCorrections2 
         Caption         =   "Calculate Binary Intensities (Output Calculated Intensities for All Matrix Corrections and MAC Files, All Lines, Multiple Files)"
         HelpContextID   =   140
      End
      Begin VB.Menu menuAnalyticalSeparator5 
         Caption         =   "-"
      End
      Begin VB.Menu menuAnalyticalCalculateStandardConcentrations 
         Caption         =   "Calculate Standard Concentrations (Output Calculated Concentrations and Errors)  (Test with CALCZAF2.DAT)"
         HelpContextID   =   603
      End
      Begin VB.Menu menuAnalyticalCalculateStandardConcentrationsAllCorrections 
         Caption         =   "Calculate Standard Concentrations for All Matrix Corrections and MAC Files, Single Line, Single File) (Test with CALCZAF2.DAT)"
         HelpContextID   =   604
      End
      Begin VB.Menu menuAnalyticalSeparator6 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu menuAnalyticalCalculateBinaryIntensities1 
         Caption         =   "Calculate Binary Intensities (Output Corrected Atomic First Approximation)"
         HelpContextID   =   141
         Visible         =   0   'False
      End
      Begin VB.Menu menuAnalyticalCalculateBinaryIntensities2 
         Caption         =   "Calculate Binary Intensities (Output Corrected Mass First Approximation)"
         HelpContextID   =   142
         Visible         =   0   'False
      End
      Begin VB.Menu menuAnalyticalCalculateBinaryIntensities3 
         Caption         =   "Calculate Binary Intensities (Output Corrected Electron First Approximation)"
         HelpContextID   =   143
         Visible         =   0   'False
      End
      Begin VB.Menu menuAnalyticalCalculateFirstApproximations1 
         Caption         =   "Calculate First Approximations Only (Atomic Fraction)"
         HelpContextID   =   144
         Visible         =   0   'False
      End
      Begin VB.Menu menuAnalyticalCalculateFirstApproximations2 
         Caption         =   "Calculate First Approximations Only (Mass Fraction)"
         HelpContextID   =   145
         Visible         =   0   'False
      End
      Begin VB.Menu menuAnalyticalCalculateFirstApproximations3 
         Caption         =   "Calculate First Approximations Only (Electron Fraction)"
         HelpContextID   =   146
         Visible         =   0   'False
      End
   End
   Begin VB.Menu menuRun 
      Caption         =   "&Run"
      HelpContextID   =   147
      Begin VB.Menu menuRunListStandardCompositions 
         Caption         =   "&List Standard Compositions"
         HelpContextID   =   148
      End
      Begin VB.Menu menuRunListCurrentMACs 
         Caption         =   "List Current MACs"
         HelpContextID   =   149
      End
      Begin VB.Menu menuRunListCurrentAlphas 
         Caption         =   "List Current Alpha Factors"
      End
      Begin VB.Menu menuRunListAnalysisParameters 
         Caption         =   "List Analysis Parameters"
         HelpContextID   =   150
      End
      Begin VB.Menu menuRunSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu menuRunModelDetectionLimits 
         Caption         =   "&Model Detection Limits"
         HelpContextID   =   151
      End
      Begin VB.Menu menuRunCalculateElectronXrayRanges 
         Caption         =   "&Calculate Electron and Xray Ranges"
         HelpContextID   =   775
      End
      Begin VB.Menu menuRunCalculateTemperatureRise 
         Caption         =   "Calculate Sample Temperature Rise"
         HelpContextID   =   776
      End
   End
   Begin VB.Menu menuOutput 
      Caption         =   "&Output"
      HelpContextID   =   153
      Begin VB.Menu menuOutputLogWindow 
         Caption         =   "Log Window Font"
         HelpContextID   =   154
      End
      Begin VB.Menu menuOutputDebugMode 
         Caption         =   "Debug Mode"
         HelpContextID   =   155
      End
      Begin VB.Menu menuOutputVerboseMode 
         Caption         =   "Verbose Mode"
         HelpContextID   =   687
      End
      Begin VB.Menu menuOutputSeparator0 
         Caption         =   "-"
      End
      Begin VB.Menu menuOutputUseAutomaticFormatForResults 
         Caption         =   "Use Automatic Format For Results"
      End
      Begin VB.Menu menuOutputZAFEquationMode 
         Caption         =   "ZAF Equation Mode (1/A, 1+F, 1/S)"
         HelpContextID   =   777
      End
      Begin VB.Menu menuOutputSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu menuOutputExtendedFormat 
         Caption         =   "Extended Format"
         HelpContextID   =   156
      End
      Begin VB.Menu menuOutputSaveToDiskLog 
         Caption         =   "Save To Disk Log"
         HelpContextID   =   157
      End
      Begin VB.Menu menuViewDiskLog 
         Caption         =   "View Disk Log"
         HelpContextID   =   158
      End
      Begin VB.Menu menuOutputSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu menuOutputOpenLinkToExcel 
         Caption         =   "Open Link To Excel"
         HelpContextID   =   159
      End
      Begin VB.Menu menuOutputCloseLinkToExcel 
         Caption         =   "Close Link To Excel"
         Checked         =   -1  'True
         HelpContextID   =   160
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "&Help"
      HelpContextID   =   161
      Begin VB.Menu menuHelpAboutCalcZAF 
         Caption         =   "&About CalcZAF"
         HelpContextID   =   162
      End
      Begin VB.Menu menuHelpOnCalcZAF 
         Caption         =   "&Help on CalcZAF"
         HelpContextID   =   163
      End
      Begin VB.Menu menuHelpGettingStartedWithCalcZAF 
         Caption         =   "&Getting Started With CalcZAF"
      End
      Begin VB.Menu menuHelpSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu menuHelpUpdateCalcZAF 
         Caption         =   "&Update CalcZAF"
         HelpContextID   =   722
      End
      Begin VB.Menu menuHelpSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu menuHelpProbeSoftwareOnTheWeb 
         Caption         =   "&Probe Software On The Web"
         HelpContextID   =   778
      End
      Begin VB.Menu menuHelpProbeSoftwareUserForum 
         Caption         =   "&Connect To Probe Software User Forum"
      End
   End
End
Attribute VB_Name = "FormMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2022 by John J. Donovan
Option Explicit

Private Sub Form_Activate()
Static initialized As Integer
If Not DebugMode Then On Error Resume Next

' Initialize the global variables
If initialized = False Then
initialized = True

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

' Check if the standard database needs to be updated
Call StandardUpdateMDBFile(StandardDataFile$)
If ierror Then Exit Sub

' Debug mode is default for CalcZAF
RealTimeMode = False
DebugMode% = True
FormMAIN.menuOutputDebugMode.Checked = True

' Initialize the log window
FormMAIN.TimerLogWindow.Interval = Int(LogWindowInterval! * 1000)

' Init samples
Call CalcZAFInit
If ierror Then Exit Sub

' Show the element grid
FormZAF.Show vbModeless
Call CalcZAFLoad
If ierror Then Exit Sub

'Call CalcZAFConvertPouchouCSV          ' test code
'Call EditConvertCSVToText(FormMAIN)    ' test code
'Call EditConvertFFAST2Dat              ' test code

Call InitWindow(Int(2), MDBUserName$, Me)           ' to ensure default position gets saved
End If

End Sub

Private Sub Form_Load()
' Load FormMAIN for program CALCZAF
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)

' Load form and application icon
Call MiscLoadIcon(FormMAIN)

' Check if program is already running
If app.PrevInstance Then
msg$ = "CALCZAF is already running, click OK, then type <ctrl> <esc> for the Task Manager and select CALCZAF.EXE from the Task List"
MsgBox msg$, vbOKOnly + vbExclamation, "CALCZAF"
End
End If

' Help file
FormMAIN.HelpContextID = IOGetHelpContextID("FormMAIN")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Not DebugMode Then On Error Resume Next
Call ExcelCloseSpreadsheet(vbNullString, FormMAIN)
If ierror Then Exit Sub
Unload FormZAF
End Sub

Private Sub Form_Resize()
If Not DebugMode Then On Error Resume Next
Dim tempsize As Integer

' Make text box (Log Window) full size of window
FormMAIN.TextLog.Left = 0
If FormMAIN.ScaleWidth > 0 Then FormMAIN.TextLog.Width = FormMAIN.ScaleWidth

tempsize% = FormMAIN.ScaleHeight - FormMAIN.TextLog.Top - FormMAIN.StatusBarAuto.Height
If tempsize% > 0 Then FormMAIN.TextLog.Height = tempsize%
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
' Disable timer
FormMAIN.TimerLogWindow.Enabled = False
End
End Sub

Private Sub menuAnalyticalAlphaFactors_Click()
If Not DebugMode Then On Error Resume Next
Call CalcZAFPlotAlphas
If ierror Then Exit Sub
End Sub

Private Sub menuAnalyticalBinaryCalculationOptions_Click()
If Not DebugMode Then On Error Resume Next

' Load the options
Call CalcZAFBinaryLoad
If ierror Then Exit Sub

FormBINARY.Show vbModal
DoEvents
End Sub

Private Sub menuAnalyticalCalculateBinaryIntensities_Click()
If Not DebugMode Then On Error Resume Next

' Close existing data file if necessary
Call CalcZAFImportClose
If ierror Then Exit Sub
FormMAIN.TextLog.Text = vbNullString

' Load the histogram options
Call CalcZAFHistogramLoad
If ierror Then Exit Sub

' Calculate binary compositions
Call CalcZAFBinary(Int(0), FormMAIN)
If ierror Then Exit Sub
End Sub

Private Sub menuAnalyticalCalculateBinaryIntensities1_Click()
If Not DebugMode Then On Error Resume Next

' Close existing data file if necessary
Call CalcZAFImportClose
If ierror Then Exit Sub
FormMAIN.TextLog.Text = vbNullString

' Load the options
Call CalcZAFHistogramLoad
If ierror Then Exit Sub

' Calculate atomic first approximation
Call CalcZAFBinary(Int(1), FormMAIN)
If ierror Then Exit Sub
End Sub

Private Sub menuAnalyticalCalculateBinaryIntensities2_Click()
If Not DebugMode Then On Error Resume Next

' Close existing data file if necessary
Call CalcZAFImportClose
If ierror Then Exit Sub
FormMAIN.TextLog.Text = vbNullString

' Load the options
Call CalcZAFHistogramLoad
If ierror Then Exit Sub

' Calculate mass first approximation
Call CalcZAFBinary(Int(2), FormMAIN)
If ierror Then Exit Sub
End Sub

Private Sub menuAnalyticalCalculateBinaryIntensities3_Click()
If Not DebugMode Then On Error Resume Next

' Close existing data file if necessary
Call CalcZAFImportClose
If ierror Then Exit Sub
FormMAIN.TextLog.Text = vbNullString

' Load the options
Call CalcZAFHistogramLoad
If ierror Then Exit Sub

' Calculate electron first approximation
Call CalcZAFBinary(Int(3), FormMAIN)
If ierror Then Exit Sub
End Sub

Private Sub menuAnalyticalCalculateBinaryIntensitiesAllCorrections_Click()
If Not DebugMode Then On Error Resume Next
Dim itemp1 As Integer, itemp2 As Integer

' Close existing data file if necessary
Call CalcZAFImportClose
If ierror Then Exit Sub
FormMAIN.TextLog.Text = vbNullString

' Calculate standard k-factors for all corrections and MAC files (single line, single file)
itemp1% = izaf%
itemp2% = MACTypeFlag%
Call CalcZAFBinary(Int(4), FormMAIN)
izaf% = itemp1%
MACTypeFlag% = itemp2%
If ierror Then Exit Sub

' Restore original standard k-factors
Call CalcZAFUpdateAllStdKfacs
If ierror Then Exit Sub

End Sub

Private Sub menuAnalyticalCalculateBinaryIntensitiesAllCorrections2_Click()
If Not DebugMode Then On Error Resume Next
Dim itemp1 As Integer, itemp2 As Integer

' Close existing data file if necessary
Call CalcZAFImportClose
If ierror Then Exit Sub
FormMAIN.TextLog.Text = vbNullString

' Calculate standard k-factors for all corrections and MAC files (all lines, multiple files)
itemp1% = izaf%
itemp2% = MACTypeFlag%
Call CalcZAFBinary(Int(5), FormMAIN)
izaf% = itemp1%
MACTypeFlag% = itemp2%
If ierror Then Exit Sub

' Restore original standard k-factors
Call CalcZAFUpdateAllStdKfacs
If ierror Then Exit Sub

End Sub

Private Sub menuAnalyticalCalculateFirstApproximations1_Click()
If Not DebugMode Then On Error Resume Next

' Close existing data file if necessary
Call CalcZAFImportClose
If ierror Then Exit Sub
FormMAIN.TextLog.Text = vbNullString

' Load the options
Call CalcZAFHistogramLoad
If ierror Then Exit Sub

' Calculate first approximation
Call CalcZAFFirstApproximation(Int(1), FormMAIN)
If ierror Then Exit Sub
End Sub

Private Sub menuAnalyticalCalculateFirstApproximations2_Click()
If Not DebugMode Then On Error Resume Next

' Close existing data file if necessary
Call CalcZAFImportClose
If ierror Then Exit Sub
FormMAIN.TextLog.Text = vbNullString

' Load the options
Call CalcZAFHistogramLoad
If ierror Then Exit Sub

' Calculate first approximation
Call CalcZAFFirstApproximation(Int(2), FormMAIN)
If ierror Then Exit Sub
End Sub

Private Sub menuAnalyticalCalculateFirstApproximations3_Click()
If Not DebugMode Then On Error Resume Next

' Close existing data file if necessary
Call CalcZAFImportClose
If ierror Then Exit Sub
FormMAIN.TextLog.Text = vbNullString

' Load the options
Call CalcZAFHistogramLoad
If ierror Then Exit Sub

' Calculate standard k-factors
Call CalcZAFFirstApproximation(Int(3), FormMAIN)
If ierror Then Exit Sub
End Sub

Private Sub menuAnalyticalCalculateStandardConcentrations_Click()
If Not DebugMode Then On Error Resume Next

' Close existing data file if necessary
Call CalcZAFImportClose
If ierror Then Exit Sub
FormMAIN.TextLog.Text = vbNullString

' Calculate compositions of standards and compare to published values
Call CalcZAFStandard(Int(0), FormMAIN)
If ierror Then Exit Sub
End Sub

Private Sub menuAnalyticalCalculateStandardConcentrationsAllCorrections_Click()
If Not DebugMode Then On Error Resume Next
' Close existing data file if necessary
Call CalcZAFImportClose
If ierror Then Exit Sub
FormMAIN.TextLog.Text = vbNullString

' Calculate compositions of standards and compare to published values for all matrix corrections
Call CalcZAFStandard(Int(1), FormMAIN)
If ierror Then Exit Sub
End Sub

Private Sub menuAnalyticalConditions_Click()
If Not DebugMode Then On Error Resume Next

' Load the form
Call CondLoad
If ierror Then Exit Sub

' Load COND form
FormCOND.Show vbModal

' Update form (and analytical conditions)
Call CalcZAFLoad
If ierror Then Exit Sub

' Calculate standard k-factors
If CalcZAFMode% > 0 Then
Call CalcZAFUpdateAllStdKfacs
If ierror Then Exit Sub
End If

FormZAF.Show vbModeless
End Sub

Private Sub menuAnalyticalElementalFactors_Click()
If Not DebugMode Then On Error Resume Next
Call CalcZAFElementalToOxideFactors(Int(1))
If ierror Then Exit Sub
Exit Sub
End Sub

Private Sub menuAnalyticalElements_Click()
If Not DebugMode Then On Error Resume Next
FormZAF.Show vbModeless
Call CalcZAFLoad
If ierror Then Exit Sub
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

' Calculate standard k-factors
Call CalcZAFUpdateAllStdKfacs
If ierror Then Exit Sub
End Sub

Private Sub menuAnalyticalKFactorsAlphaFactors_Click()
If Not DebugMode Then On Error Resume Next
Call CalcZAFCalculateKRatiosAlphaFactors(FormMAIN)
If ierror Then Exit Sub
End Sub

Private Sub menuAnalyticalModelSeco_Click()
If Not DebugMode Then On Error Resume Next
Call CalcZAFRunStandard     ' run Standard.EXE as a separate process
If ierror Then Exit Sub
End Sub

Private Sub menuAnalyticalOxideFactors_Click()
If Not DebugMode Then On Error Resume Next
Call CalcZAFElementalToOxideFactors(Int(2))
If ierror Then Exit Sub
End Sub

Private Sub menuAnalyticalParticleandThinFilm_Click()
If Not DebugMode Then On Error Resume Next
' Get PTC options
Call CalcZAFGetPTC
If ierror Then Exit Sub
End Sub

Private Sub menuAnalyticalSecondary_Click()
If Not DebugMode Then On Error Resume Next
If Penepma08CheckPenepmaVersion%() = 12 Then
Call SecondaryLoad
If ierror Then Exit Sub
FormSECONDARY.Show vbModeless
Else
msg$ = "Penepma12 application files were not found. Please download the PENEPMA12.ZIP file and extract the files to the " & UserDataDirectory$ & " folder and check that the PENEPMA_Path, PENDBASE_Path and PENEPMA_Root strings are properly specified in the " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "menuAnayticalSecondary"
End If
End Sub

Private Sub menuAnalyticalUseConductiveCoatingCorrectionForXrayTransmission_Click()
If Not DebugMode Then On Error Resume Next
UseConductiveCoatingCorrectionForXrayTransmission = Not UseConductiveCoatingCorrectionForXrayTransmission
If Not UseConductiveCoatingCorrectionForXrayTransmission Then
FormMAIN.menuAnalyticalUseConductiveCoatingCorrectionForXrayTransmission.Checked = vbUnchecked
Else
FormMAIN.menuAnalyticalUseConductiveCoatingCorrectionForXrayTransmission.Checked = vbChecked
End If
End Sub

Private Sub menuAnalyticalUseConductiveCoatingCorrectionForElectronAbsorption_Click()
If Not DebugMode Then On Error Resume Next
UseConductiveCoatingCorrectionForElectronAbsorption = Not UseConductiveCoatingCorrectionForElectronAbsorption
If Not UseConductiveCoatingCorrectionForElectronAbsorption Then
FormMAIN.menuAnalyticalUseConductiveCoatingCorrectionForElectronAbsorption.Checked = vbUnchecked
Else
FormMAIN.menuAnalyticalUseConductiveCoatingCorrectionForElectronAbsorption.Checked = vbChecked
End If
End Sub

Private Sub menuAnalyticalZAFSelections_Click()
If Not DebugMode Then On Error Resume Next

' Update or display ZAF selections
Call GetZAFAllLoad
FormGETZAFALL.Show vbModal
If ierror Then Exit Sub

' Warn if user selects calibration curve in CalcZAF (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
If CorrectionFlag% = 5 Then
msg$ = "Calibration curve calculation is not available in CalcZAF. Re-setting to ZAF or Phi-Rho-Z calculations instead."
MsgBox msg$, vbOKOnly + vbInformation, "CalcZAF"
CorrectionFlag% = 0
End If

Call TypeZAFSelections
If ierror Then Exit Sub

' Calculate standard k-factors
Call CalcZAFUpdateAllStdKfacs
If ierror Then Exit Sub
End Sub

Private Sub menuClearAll_Click()
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

Private Sub menuFileClose_Click()
If Not DebugMode Then On Error Resume Next
Call CalcZAFImportClose
If ierror Then Exit Sub
FormMAIN.TextLog.Text = vbNullString
End Sub

Private Sub menuFileExit_Click()
If Not DebugMode Then On Error Resume Next
Unload FormMAIN
End Sub

Private Sub menuFileExport_Click()
If Not DebugMode Then On Error Resume Next
Call CalcZAFExportOpen(FormMAIN)
If ierror Then Exit Sub
Call CalcZAFExportSend
If ierror Then Exit Sub
Call CalcZAFExportClose(Int(0))
If ierror Then Exit Sub
End Sub

Private Sub menuFileOpen_Click()
If Not DebugMode Then On Error Resume Next
FormMAIN.TextLog.Text = vbNullString
Call CalcZAFImportOpen(FormMAIN)
If ierror Then Exit Sub
End Sub

Private Sub menuFileOpenAndProcess_Click()
If Not DebugMode Then On Error Resume Next
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%
FormMAIN.TextLog.Text = vbNullString
Call CalcZAFCalculateExportAll(FormMAIN)
If ierror Then Exit Sub
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

Private Sub menuFileUpdateCalcZAFSampleDataFiles_Click()
If Not DebugMode Then On Error Resume Next
Call InitFilesUserData
If ierror Then Exit Sub
End Sub

Private Sub menuHelpAboutCalcZAF_Click()
If Not DebugMode Then On Error Resume Next
FormABOUT.Show vbModal
End Sub

Private Sub menuHelpGettingStartedWithCalcZAF_Click()
If Not DebugMode Then On Error Resume Next
Call IOBrowseHTTP(ProbeSoftwareInternetBrowseMethod%, "https://probesoftware.com/smf/index.php?topic=81.0")
If ierror Then Exit Sub
End Sub

Private Sub menuHelpOnCalcZAF_Click()
If Not DebugMode Then On Error Resume Next
Call MiscFormLoadHelp(FormMAIN.HelpContextID)
If ierror Then Exit Sub
End Sub

Private Sub menuHelpProbeSoftwareOnTheWeb_Click()
If Not DebugMode Then On Error Resume Next
Call IOBrowseHTTP(ProbeSoftwareInternetBrowseMethod%, "https://probesoftware.com/index.html")
If ierror Then Exit Sub
End Sub

Private Sub menuHelpProbeSoftwareUserForum_Click()
If Not DebugMode Then On Error Resume Next
Call IOBrowseHTTP(ProbeSoftwareInternetBrowseMethod%, "https://probesoftware.com/smf/index.php")
If ierror Then Exit Sub
End Sub

Private Sub menuHelpUpdateCalcZAF_Click()
If Not DebugMode Then On Error Resume Next
FormUPDATE.Show vbModal
End Sub

Private Sub menuOutputCloseLinkToExcel_Click()
If Not DebugMode Then On Error Resume Next
Call ExcelCloseSpreadsheet(vbNullString, FormMAIN)
If ierror Then Exit Sub
FormMAIN.menuOutputOpenLinkToExcel.Checked = False
FormMAIN.menuOutputCloseLinkToExcel.Checked = True
End Sub

Private Sub menuOutputDebugMode_Click()
If Not DebugMode Then On Error Resume Next
FormMAIN.menuOutputDebugMode.Checked = Not FormMAIN.menuOutputDebugMode.Checked
DebugMode = FormMAIN.menuOutputDebugMode.Checked
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

Private Sub menuOutputOpenLinkToExcel_Click()
If Not DebugMode Then On Error Resume Next
Call ExcelCreateSpreadsheet(FormMAIN)
If ierror Then Exit Sub
FormMAIN.menuOutputOpenLinkToExcel.Checked = True
FormMAIN.menuOutputCloseLinkToExcel.Checked = False
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

Private Sub menuOutputUseAutomaticFormatForResults_Click()
If Not DebugMode Then On Error Resume Next
FormMAIN.menuOutputUseAutomaticFormatForResults.Checked = Not FormMAIN.menuOutputUseAutomaticFormatForResults.Checked
UseAutomaticFormatForResultsFlag = FormMAIN.menuOutputUseAutomaticFormatForResults.Checked
End Sub

Private Sub menuOutputVerboseMode_Click()
If Not DebugMode Then On Error Resume Next
FormMAIN.menuOutputVerboseMode.Checked = Not FormMAIN.menuOutputVerboseMode.Checked
VerboseMode = FormMAIN.menuOutputVerboseMode.Checked
End Sub

Private Sub menuOutputZAFEquationMode_Click()
If Not DebugMode Then On Error Resume Next
FormMAIN.menuOutputZAFEquationMode.Checked = Not FormMAIN.menuOutputZAFEquationMode.Checked
ZAFEquationMode = FormMAIN.menuOutputZAFEquationMode.Checked
End Sub

Private Sub menuRunCalculateElectronXrayRanges_Click()
If Not DebugMode Then On Error Resume Next
Call RangeLoad
If ierror Then Exit Sub
FormRANGE.Show vbModeless
End Sub

Private Sub menuRunCalculateTemperatureRise_Click()
If Not DebugMode Then On Error Resume Next
Call TemperatureLoad
If ierror Then Exit Sub
FormTEMPERATURE.Show vbModeless
End Sub

Private Sub menuRunListAnalysisParameters_Click()
If Not DebugMode Then On Error Resume Next
' Type out last analysis parameters
Call CalcZAFTypeAnalysis
If ierror Then Exit Sub
End Sub

Private Sub menuRunListCurrentAlphas_Click()
If Not DebugMode Then On Error Resume Next
Call CalcZAFListCurrentAlphas
If ierror Then Exit Sub
End Sub

Private Sub menuRunListCurrentMACs_Click()
If Not DebugMode Then On Error Resume Next
' List last MACs loaded
Call ZAFPrintMAC
If ierror Then Exit Sub
End Sub

Private Sub menuRunListStandardCompositions_Click()
If Not DebugMode Then On Error Resume Next
' Type out standard compositions
Call CalcZAFTypeStandards
If ierror Then Exit Sub
End Sub

Private Sub menuRunModelDetectionLimits_Click()
If Not DebugMode Then On Error Resume Next
Call DetectionLoad
If ierror Then Exit Sub
FormDETECTION.Show vbModeless
End Sub

Private Sub menuStandardAddStandardsToRun_Click()
' Add standards to the current run
If Not DebugMode Then On Error Resume Next
Call AddStdLoad
If ierror Then Exit Sub
FormADDSTD.Show vbModal
If ierror Then Exit Sub
End Sub

Private Sub menuStandardEditStandardParameters_Click()
If Not DebugMode Then On Error Resume Next
Call StandardCoatingLoad
If ierror Then Exit Sub
End Sub

Private Sub menuStandardSelectStandardDatabase_Click()
If Not DebugMode Then On Error Resume Next
Call CalcZAFSelectStandardDatabase(FormMAIN)
If ierror Then Exit Sub
End Sub

Private Sub menuStandardStandardDatabase_Click()
If Not DebugMode Then On Error Resume Next
Call CalcZAFRunStandard     ' run Standard.EXE as a separate process
If ierror Then Exit Sub
End Sub

Private Sub menuStudentstTable_Click()
If Not DebugMode Then On Error Resume Next
Call StudentCalculateTable(msg$)
If ierror Then Exit Sub
Call IOWriteLog(msg$)
End Sub

Private Sub menuViewDiskLog_Click()
If Not DebugMode Then On Error Resume Next
' View disk log file
Call IOViewLog
If ierror Then Exit Sub
End Sub

Private Sub menuXrayCalculateSpectrometerPosition_Click()
If Not DebugMode Then On Error Resume Next
Call CalcSpecLoad
If ierror Then Exit Sub
End Sub

Private Sub menuXrayConvertELEMINFODAT_Click()
If Not DebugMode Then On Error Resume Next
Call EditConvertElemInfoDat
If ierror Then Exit Sub
End Sub

Private Sub menuXrayConvertMACMATDAT_Click()
If Not DebugMode Then On Error Resume Next
Call EditConvertMACMatDat
If ierror Then Exit Sub
End Sub

Private Sub menuXrayCreateNewFFASTMACTable_Click()
If Not DebugMode Then On Error Resume Next
Call EditMakeNewMACTable(Int(4))
If ierror Then Exit Sub
End Sub

Private Sub menuXrayCreateNewMAC30MACTable_Click()
If Not DebugMode Then On Error Resume Next
Call EditMakeNewMACTable(Int(2))
If ierror Then Exit Sub
End Sub

Private Sub menuXrayCreateNewMACJTAMACTable_Click()
If Not DebugMode Then On Error Resume Next
Call EditMakeNewMACTable(Int(3))
If ierror Then Exit Sub
End Sub

Private Sub menuXrayCreateNewMcMasterMACTable_Click()
If Not DebugMode Then On Error Resume Next
Call EditMakeNewMACTable(Int(1))
If ierror Then Exit Sub
End Sub

Private Sub menuXrayCreateNewUSERMACTable_Click()
If Not DebugMode Then On Error Resume Next
Call EditMakeNewMACTable(Int(5))
If ierror Then Exit Sub
End Sub

Private Sub menuXrayCreateNewXrayDatabase_Click()
If Not DebugMode Then On Error Resume Next
' Create a new XRAY.MDB file (requires XRAY.ALL)
Call XrayOpenNewMDB
If ierror Then Exit Sub
End Sub

Private Sub menuXrayDisplayMACEmitterAbsorber_Click()
If Not DebugMode Then On Error Resume Next
Call EditGetMACEmitterAbsorber
If ierror Then Exit Sub
End Sub

Private Sub menuXrayEdgeTable_Click()
If Not DebugMode Then On Error Resume Next
' Obtain an edge table
Call XrayGetTable(Int(2))
If ierror Then Exit Sub
End Sub

Private Sub menuXrayEditMACTable_Click()
If Not DebugMode Then On Error Resume Next
Call EditMACLoad
If ierror Then Exit Sub
FormEDITMAC.Show vbModal
End Sub

Private Sub menuXrayEditXedgeTable_Click()
If Not DebugMode Then On Error Resume Next
Call EditXrayLoad(Int(2))
If ierror Then Exit Sub
FormEDITXRAY.Show vbModal
End Sub

Private Sub menuXrayEditXflurTable_Click()
If Not DebugMode Then On Error Resume Next
Call EditXrayLoad(Int(3))
If ierror Then Exit Sub
FormEDITXRAY.Show vbModal
End Sub

Private Sub menuXrayEditXrayTable_Click()
If Not DebugMode Then On Error Resume Next
Call EditXrayLoad(Int(1))
If ierror Then Exit Sub
FormEDITXRAY.Show vbModal
End Sub

Private Sub menuXrayEmissionTable_Click()
If Not DebugMode Then On Error Resume Next
' Obtain an emission table
Call XrayGetTable(Int(1))
If ierror Then Exit Sub
End Sub

Private Sub menuXrayEmissionTable2_Click()
If Not DebugMode Then On Error Resume Next
' Obtain an emission table for additional x-rays
Call XrayGetTable(Int(4))
If ierror Then Exit Sub
End Sub

Private Sub menuXrayFluorescentYieldTable_Click()
If Not DebugMode Then On Error Resume Next
' Obtain an fluorescent yield table
Call XrayGetTable(Int(3))
If ierror Then Exit Sub
End Sub

Private Sub menuXrayFluorescentYieldTable2_Click()
If Not DebugMode Then On Error Resume Next
' Obtain an emission table for additional x-rays
Call XrayGetTable(Int(6))
If ierror Then Exit Sub
End Sub

Private Sub menuXrayMACTable_Click()
If Not DebugMode Then On Error Resume Next
' Obtain a MAC table
Call XrayGetTable(Int(7))
If ierror Then Exit Sub
End Sub

Private Sub menuXrayMACTable2_Click()
If Not DebugMode Then On Error Resume Next
' Obtain a MAC table (additional lines)
Call XrayGetTable(Int(9))
If ierror Then Exit Sub
End Sub

Private Sub menuXrayMACTableComplete_Click()
If Not DebugMode Then On Error Resume Next
' Obtain a MAC table (complete)
Call XrayGetTable(Int(8))
If ierror Then Exit Sub
End Sub

Private Sub menuXrayMACTableComplete2_Click()
If Not DebugMode Then On Error Resume Next
' Obtain a MAC table (complete) (additional lines)
Call XrayGetTable(Int(10))
If ierror Then Exit Sub
End Sub

Private Sub menuXrayOutputExistingUSERMACTable_Click()
If Not DebugMode Then On Error Resume Next
Call EditOutputUserMACFile(FormMAIN)
If ierror Then Exit Sub
End Sub

Private Sub menuXrayUpdateUSERMACTable_Click()
If Not DebugMode Then On Error Resume Next
Call EditUpdateUserMACFile(FormMAIN)
If ierror Then Exit Sub
End Sub

Private Sub menuXrayUpdateXEdgeTable_Click()
If Not DebugMode Then On Error Resume Next
Call EditUpdateXFiles(Int(2), FormMAIN)
If ierror Then Exit Sub
End Sub

Private Sub menuXrayUpdateXFlurTable_Click()
If Not DebugMode Then On Error Resume Next
Call EditUpdateXFiles(Int(3), FormMAIN)
If ierror Then Exit Sub
End Sub

Private Sub menuXrayUpdateXLineTable_Click()
If Not DebugMode Then On Error Resume Next
Call EditUpdateXFiles(Int(1), FormMAIN)
If ierror Then Exit Sub
End Sub

Private Sub menuXrayXrayDatabase_Click()
If Not DebugMode Then On Error Resume Next
' Loads the xray database
Call XrayGetDatabase
If ierror Then Exit Sub
End Sub

Private Sub menuXrayConvertTextToData_Click()
If Not DebugMode Then On Error Resume Next
Call EditConvertTextToDAT(FormMAIN)
If ierror Then Exit Sub
End Sub

Private Sub menuXrayConvertDataToText_Click()
If Not DebugMode Then On Error Resume Next
Call EditConvertDATToText(FormMAIN)
If ierror Then Exit Sub
End Sub

Private Sub menuXrayUpdateEdgeLineFlurFiles_Click()
If Not DebugMode Then On Error Resume Next
Call EditUpdateEdgeLineFlurFiles
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

