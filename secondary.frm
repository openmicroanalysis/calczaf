VERSION 5.00
Begin VB.Form FormSECONDARY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Perform Correction For Secondary Fluorescence From Boundary"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   13440
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameUpdateBoundary 
      Caption         =   "Update Boundary Position"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   9600
      TabIndex        =   53
      Top             =   720
      Visible         =   0   'False
      Width           =   9375
      Begin VB.CommandButton CommandUpdatePositionCoordinateAngle 
         Caption         =   "Update Position of Boundary Coordinate (and angle)"
         Height          =   735
         Left            =   240
         TabIndex        =   56
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton CommandUpdatePositionCoordinatePair1 
         Caption         =   "Update Positions of Boundary Coordinate (first pair)"
         Height          =   735
         Left            =   2880
         TabIndex        =   55
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton CommandUpdatePositionCoordinatePair2 
         Caption         =   "Update Positions of Boundary Coordinate (second pair)"
         Height          =   735
         Left            =   5160
         TabIndex        =   54
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label LabelUpdatePositions 
         Alignment       =   2  'Center
         Caption         =   "Adjust the stage for the boundary position and click one of the update position buttons"
         Height          =   855
         Left            =   7560
         TabIndex        =   57
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   735
      Left            =   7440
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   52
      Top             =   6000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   735
      Left            =   6000
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   51
      Top             =   6000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton CommandPrintImage 
      Caption         =   "Print Image"
      Height          =   375
      Left            =   4800
      TabIndex        =   50
      Top             =   8400
      Width           =   2175
   End
   Begin VB.CommandButton CommandHelp 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Click this button to get detailed help from our on-line user forum"
      Top             =   120
      Width           =   1455
   End
   Begin VB.CheckBox CheckUseSecondaryFluorescenceCorrection 
      Caption         =   "Perform Boundary Correction (invisible)"
      Height          =   255
      Left            =   9960
      TabIndex        =   47
      Top             =   480
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   735
      Left            =   6720
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   46
      Top             =   3360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton CommandCopyToClipboard 
      Caption         =   "Copy To Clipboard"
      Height          =   375
      Left            =   7200
      TabIndex        =   45
      ToolTipText     =   "Copy the above image (with graphics onjects) to the system clipboard"
      Top             =   8400
      Width           =   2175
   End
   Begin VB.TextBox TextHFW 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7320
      TabIndex        =   44
      Top             =   8880
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8880
      Top             =   8760
   End
   Begin VB.Frame Frame4 
      Caption         =   "Secondary Fluorescence Boundary Correction Method"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Width           =   9255
      Begin VB.CommandButton CommandBrowseForCouple 
         BackColor       =   &H0080FFFF&
         Caption         =   "Browse For SF Couple"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   $"Secondary.frx":0000
         Top             =   240
         Width           =   2295
      End
      Begin VB.OptionButton OptionCorrectionMethod 
         Caption         =   "Calculate Using Binary Composition K-Ratios From Matrix and Boundary Databases"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Visible         =   0   'False
         Width           =   8895
      End
      Begin VB.OptionButton OptionCorrectionMethod 
         Caption         =   "Calculate Using K-Ratios From Previously Calculated PAR File Couple"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   6615
      End
      Begin VB.Label LabelKratiosDATFile 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   720
         Width           =   8535
      End
   End
   Begin VB.Frame FrameImage 
      Caption         =   "Distance To Boundary Method"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   6615
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   4455
      Begin VB.TextBox TextY2StageCoordinate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   38
         Top             =   4680
         Width           =   1095
      End
      Begin VB.TextBox TextX2StageCoordinate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   36
         Top             =   4320
         Width           =   1095
      End
      Begin VB.TextBox TextY1StageCoordinate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   34
         Top             =   3840
         Width           =   1095
      End
      Begin VB.TextBox TextX1StageCoordinate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   32
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox TextBoundaryAngle 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   30
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox TextYStageCoordinate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   28
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox TextXStageCoordinate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   26
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton CommandBrowse 
         BackColor       =   &H0080FFFF&
         Caption         =   "Browse For Image"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   5400
         Width           =   2535
      End
      Begin VB.OptionButton OptionDistanceMethod 
         Caption         =   "Specify Graphical Boundary"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   23
         ToolTipText     =   "Specify the boundary by drawing a line on a stage calibrated image"
         Top             =   5160
         Width           =   2775
      End
      Begin VB.OptionButton OptionDistanceMethod 
         Caption         =   "Specify X,Y Coordinate Pair"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   22
         ToolTipText     =   "Specify boundary as a pair of X,Y coordinates"
         Top             =   3120
         Width           =   3255
      End
      Begin VB.OptionButton OptionDistanceMethod 
         Caption         =   "Specify X,Y Coordinate and Angle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   16
         ToolTipText     =   "Specify the boundary as a coordinate and an angle (0 to 180 where 0 degrees equals north-south, 90 degrees equals east-west)"
         Top             =   1320
         Width           =   3615
      End
      Begin VB.OptionButton OptionDistanceMethod 
         Caption         =   "Specify Fixed Distance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   15
         ToolTipText     =   "Specify the boundary as a fixed distance for all calculations"
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox TextSpecifiedDistance 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label LabelImageBMPFile 
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   120
         TabIndex        =   25
         Top             =   5880
         Width           =   4215
      End
      Begin VB.Label Label12 
         Caption         =   "Y2 Stage Coordinate"
         Height          =   255
         Left            =   600
         TabIndex        =   39
         Top             =   4680
         Width           =   2295
      End
      Begin VB.Label Label11 
         Caption         =   "X2 Stage Coordinate"
         Height          =   255
         Left            =   600
         TabIndex        =   37
         Top             =   4320
         Width           =   2295
      End
      Begin VB.Label Label10 
         Caption         =   "Y1 Stage Coordinate"
         Height          =   255
         Left            =   600
         TabIndex        =   35
         Top             =   3840
         Width           =   2295
      End
      Begin VB.Label Label9 
         Caption         =   "X1 Stage Coordinate"
         Height          =   255
         Left            =   600
         TabIndex        =   33
         Top             =   3480
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "Boundary Angle (0 to 180)"
         Height          =   255
         Left            =   600
         TabIndex        =   31
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "Y Stage Coordinate"
         Height          =   255
         Left            =   600
         TabIndex        =   29
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "X Stage Coordinate"
         Height          =   255
         Left            =   600
         TabIndex        =   27
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Constant Distance (um)"
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   720
         Width           =   2295
      End
   End
   Begin VB.CommandButton CommandClose 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Boundary Material (Mat. B)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4215
      Left            =   9600
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton CommandCompositionStandard 
         Caption         =   "Enter Composition as Standard"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   3495
      End
      Begin VB.CommandButton CommandCompositionWeight 
         Caption         =   "Enter Composition as Weight String"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1440
         Width           =   3495
      End
      Begin VB.CommandButton CommandCompositionAtom 
         Caption         =   "Enter Composition as Atom String"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox TextMaterialBComposition 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   2040
         Width           =   3255
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Define the composition of Material B by by formula, weight or standard composition."
         Height          =   615
         Left            =   600
         TabIndex        =   5
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Perform Boundary Correction on Mat A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3855
      Left            =   9600
      TabIndex        =   0
      Top             =   5280
      Width           =   3735
      Begin VB.CommandButton CommandCalculateExportAll 
         BackColor       =   &H0080FFFF&
         Caption         =   "Open Input Data File and Calculate/Export All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2880
         Width           =   3015
      End
      Begin VB.CommandButton CommandCalculateCurrent 
         BackColor       =   &H0080FFFF&
         Caption         =   "Calculate Current Sample Composition"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "The composition of Material A is defined by the current intensities or k-ratios in the CalcZAF ""Calculate ZAF Corrections"" window."
         Height          =   855
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "The composition of Material A is defined by the CalcZAF input data file."
         Height          =   495
         Left            =   480
         TabIndex        =   40
         Top             =   2400
         Width           =   2775
      End
   End
   Begin VB.Label LabelCursorPosition 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4800
      TabIndex        =   48
      Top             =   2400
      Width           =   4575
   End
   Begin VB.Label Label14 
      Caption         =   "Horizontal Field Width (um)"
      Height          =   255
      Left            =   5280
      TabIndex        =   43
      Top             =   8880
      Width           =   2055
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4575
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   4575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Keep image control square in size"
      Height          =   615
      Left            =   6120
      TabIndex        =   42
      Top             =   5040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label LabelBoundaryCoordinates 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4800
      TabIndex        =   41
      Top             =   2640
      Width           =   4575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   $"Secondary.frx":00DC
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   120
      Width           =   8415
   End
End
Attribute VB_Name = "FormSECONDARY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2018 by John J. Donovan
Option Explicit

Dim ImageX1 As Single, ImageY1 As Single
Dim ImageX2 As Single, ImageY2 As Single

Private Sub CommandBrowse_Click()
If Not DebugMode Then On Error Resume Next
Call SecondaryBrowseFile(Int(1), FormSECONDARY)
If ierror Then Exit Sub
FormSECONDARY.OptionDistanceMethod(3).Value = True
End Sub

Private Sub CommandBrowseForCouple_Click()
If Not DebugMode Then On Error Resume Next
Call SecondaryBrowseFile(Int(0), FormSECONDARY)
If ierror Then
msg$ = "There was an error reading the K-ratio2.dat file. Make sure the DAT file is in a folder which is named so it contains the beam incident, boundary and standard materials, and element atomic number and x-ray line (e.g., 15_SiO2_TiO2_TiO2_22_1)."
MsgBox msg$, vbOKOnly + vbExclamation, "FormSECONDARY"
Exit Sub
End If
End Sub

Private Sub CommandCalculateCurrent_Click()
If Not DebugMode Then On Error Resume Next
Call SecondarySave
If ierror Then Exit Sub
CalculateAllMatrixCorrections = False
UseSecondaryBoundaryFluorescenceCorrectionFlag = True
Call SecondaryInit1
If ierror Then Exit Sub
Call CalcZAFCalculate
UseSecondaryBoundaryFluorescenceCorrectionFlag = False
If ierror Then Exit Sub
End Sub

Private Sub CommandCalculateExportAll_Click()
If Not DebugMode Then On Error Resume Next
Call SecondarySave
If ierror Then Exit Sub
CalculateAllMatrixCorrections = False
UseSecondaryBoundaryFluorescenceCorrectionFlag = True
Call SecondaryInit1
If ierror Then Exit Sub
Call CalcZAFCalculateExportAll(FormSECONDARY)
UseSecondaryBoundaryFluorescenceCorrectionFlag = False
If ierror Then Exit Sub
End Sub

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
Call SecondarySave
If ierror Then Exit Sub
Unload FormSECONDARY
End Sub

Private Sub CommandCompositionAtom_Click()
If Not DebugMode Then On Error Resume Next
Call SecondaryGetComposition(Int(1))
If ierror Then Exit Sub
End Sub

Private Sub CommandCompositionStandard_Click()
If Not DebugMode Then On Error Resume Next
Call SecondaryGetComposition(Int(3))
If ierror Then Exit Sub
End Sub

Private Sub CommandCompositionWeight_Click()
If Not DebugMode Then On Error Resume Next
Call SecondaryGetComposition(Int(2))
If ierror Then Exit Sub
End Sub

Private Sub CommandCopyToClipboard_Click()
If Not DebugMode Then On Error Resume Next
' Clipboard (use special function to save graphics methods)
Call SecondaryCopyToClipboard
If ierror Then Exit Sub
End Sub

Private Sub CommandHelp_Click()
If Not DebugMode Then On Error Resume Next
Call IOBrowseHTTP(ProbeSoftwareInternetBrowseMethod%, "http://probesoftware.com/smf/index.php?topic=58.msg223#msg223")
If ierror Then Exit Sub
End Sub

Private Sub CommandPrintImage_Click()
If Not DebugMode Then On Error Resume Next
' Print image and graphical elements to printer
Call SecondaryPrintImage(FormSECONDARY)
If ierror Then Exit Sub
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormSECONDARY)
HelpContextID = IOGetHelpContextID("FormSECONDARY")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not DebugMode Then On Error Resume Next
ImageX1! = X!    ' store for boundary draw
ImageY1! = Y!    ' store for boundary draw
Call SecondaryGetBoundary(Int(1), ImageX1!, ImageY1!, ImageX2!, ImageY2!, FormSECONDARY)
If ierror Then Exit Sub
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not DebugMode Then On Error Resume Next
' Update the stage cursor
Call SecondaryUpdateCursor(X!, Y!, FormSECONDARY)
If ierror Then Exit Sub
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not DebugMode Then On Error Resume Next
ImageX2! = X!
ImageY2! = Y!
Call SecondaryGetBoundary(Int(2), ImageX1!, ImageY1!, ImageX2!, ImageY2!, FormSECONDARY)
If ierror Then Exit Sub
End Sub

Private Sub OptionCorrectionMethod_Click(Index As Integer)
If Not DebugMode Then On Error Resume Next
If Index% = 0 Then
FormSECONDARY.CommandBrowseForCouple.Enabled = True
FormSECONDARY.CommandCompositionAtom.Enabled = False
FormSECONDARY.CommandCompositionWeight.Enabled = False
FormSECONDARY.CommandCompositionStandard.Enabled = False
FormSECONDARY.TextMaterialBComposition.Enabled = False
Else
FormSECONDARY.CommandBrowseForCouple.Enabled = False
FormSECONDARY.CommandCompositionAtom.Enabled = True
FormSECONDARY.CommandCompositionWeight.Enabled = True
FormSECONDARY.CommandCompositionStandard.Enabled = True
FormSECONDARY.TextMaterialBComposition.Enabled = True
End If
End Sub

Private Sub OptionDistanceMethod_Click(Index As Integer)
If Not DebugMode Then On Error Resume Next
Call SecondaryLoadMode(Index%)
If ierror Then Exit Sub
End Sub

Private Sub TextBoundaryAngle_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextHFW_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextSpecifiedDistance_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextX1StageCoordinate_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextX2StageCoordinate_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextXStageCoordinate_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextY1StageCoordinate_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextY2StageCoordinate_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextYStageCoordinate_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub Timer1_Timer()
If Not DebugMode Then On Error Resume Next
Call SecondaryUpdateBoundary
If ierror Then Exit Sub
Call SecondaryDrawPoints(FormSECONDARY)
If ierror Then Exit Sub
End Sub
