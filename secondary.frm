VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form FormSECONDARY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Perform Correction For Secondary Fluorescence From Boundary"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   13455
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Click Element Row to Edit Boundary Correction Parameters"
      ClipControls    =   0   'False
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
      TabIndex        =   42
      Top             =   720
      Width           =   13215
      Begin MSFlexGridLib.MSFlexGrid GridElementList 
         Height          =   975
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   1720
         _Version        =   393216
         Rows            =   73
         Cols            =   7
         FixedCols       =   5
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
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
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Click this button to get detailed help from our on-line user forum"
      Top             =   8640
      Width           =   1335
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   735
      Left            =   7440
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   39
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
      TabIndex        =   38
      Top             =   6000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton CommandPrintImage 
      Caption         =   "Print Image"
      Height          =   375
      Left            =   4800
      TabIndex        =   37
      Top             =   8400
      Width           =   2175
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   735
      Left            =   6720
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   35
      Top             =   3360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton CommandCopyToClipboard 
      Caption         =   "Copy To Clipboard"
      Height          =   375
      Left            =   7200
      TabIndex        =   34
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
      TabIndex        =   33
      Top             =   8880
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8880
      Top             =   8760
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
      Height          =   6735
      Left            =   120
      TabIndex        =   5
      Top             =   2400
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
         TabIndex        =   27
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
         TabIndex        =   25
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
         TabIndex        =   23
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
         TabIndex        =   21
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
         TabIndex        =   19
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
         TabIndex        =   17
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
         TabIndex        =   15
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
         TabIndex        =   13
         Top             =   5520
         Width           =   2295
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
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label LabelImageBMPFile 
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   6000
         Width           =   4215
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "Use mouse to draw the boundary!"
         Height          =   735
         Left            =   3240
         TabIndex        =   45
         Top             =   5280
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Y2 Stage Coordinate"
         Height          =   255
         Left            =   600
         TabIndex        =   28
         Top             =   4680
         Width           =   2295
      End
      Begin VB.Label Label11 
         Caption         =   "X2 Stage Coordinate"
         Height          =   255
         Left            =   600
         TabIndex        =   26
         Top             =   4320
         Width           =   2295
      End
      Begin VB.Label Label10 
         Caption         =   "Y1 Stage Coordinate"
         Height          =   255
         Left            =   600
         TabIndex        =   24
         Top             =   3840
         Width           =   2295
      End
      Begin VB.Label Label9 
         Caption         =   "X1 Stage Coordinate"
         Height          =   255
         Left            =   600
         TabIndex        =   22
         Top             =   3480
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "Boundary Angle (0 to 180)"
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "Y Stage Coordinate"
         Height          =   255
         Left            =   600
         TabIndex        =   18
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "X Stage Coordinate"
         Height          =   255
         Left            =   600
         TabIndex        =   16
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Constant Distance (um)"
         Height          =   255
         Left            =   600
         TabIndex        =   6
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
      Height          =   495
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Perform Boundary Correction"
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
      Height          =   6135
      Left            =   9600
      TabIndex        =   0
      Top             =   2400
      Width           =   3735
      Begin VB.CheckBox CheckUseSecondaryBoundaryFluorescenceCorrection 
         Caption         =   "Use Secondary Boundary Fluorescence Correction"
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
         Left            =   480
         TabIndex        =   44
         TabStop         =   0   'False
         ToolTipText     =   "Use the secondary fluorescence correction for boundary effects"
         Top             =   480
         Value           =   1  'Checked
         Width           =   2655
      End
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
         TabIndex        =   10
         Top             =   5280
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
         TabIndex        =   3
         Top             =   2400
         Width           =   3015
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "OR"
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
         Left            =   1560
         TabIndex        =   40
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   $"Secondary.frx":0000
         Height          =   975
         Left            =   480
         TabIndex        =   1
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "The composition of the beam incident material (and optionally stage coordinates) are defined by the CalcZAF input data file"
         Height          =   855
         Left            =   480
         TabIndex        =   29
         Top             =   4320
         Width           =   2775
      End
   End
   Begin VB.Label LabelCursorPosition 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4800
      TabIndex        =   36
      Top             =   2400
      Width           =   4575
   End
   Begin VB.Label Label14 
      Caption         =   "Horizontal Field Width (um)"
      Height          =   255
      Left            =   5280
      TabIndex        =   32
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
      TabIndex        =   31
      Top             =   5040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label LabelBoundaryCoordinates 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4800
      TabIndex        =   30
      Top             =   2640
      Width           =   4575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   $"Secondary.frx":009E
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   9855
   End
End
Attribute VB_Name = "FormSECONDARY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2025 by John J. Donovan
Option Explicit

Dim ImageX1 As Single, ImageY1 As Single
Dim ImageX2 As Single, ImageY2 As Single

Private Sub CommandBrowse_Click()
If Not DebugMode Then On Error Resume Next
Call SecondaryBrowseFile(Int(1), FormSECONDARY)
If ierror Then Exit Sub
FormSECONDARY.OptionDistanceMethod(3).Value = True
End Sub

Private Sub CommandCalculateCurrent_Click()
If Not DebugMode Then On Error Resume Next
Call SecondarySave
If ierror Then Exit Sub
If FormSECONDARY.CheckUseSecondaryBoundaryFluorescenceCorrection.Value = vbChecked Then
UseSecondaryBoundaryFluorescenceCorrectionFlag = True
Call SecondaryInit1
If ierror Then Exit Sub
Else
UseSecondaryBoundaryFluorescenceCorrectionFlag = False
End If
CalculateAllMatrixCorrections = False
Call CalcZAFCalculate
If ierror Then Exit Sub
End Sub

Private Sub CommandCalculateExportAll_Click()
If Not DebugMode Then On Error Resume Next
Call SecondarySave
If ierror Then Exit Sub
If FormSECONDARY.CheckUseSecondaryBoundaryFluorescenceCorrection.Value = vbChecked Then
UseSecondaryBoundaryFluorescenceCorrectionFlag = True
Call SecondaryInit1
If ierror Then Exit Sub
Else
UseSecondaryBoundaryFluorescenceCorrectionFlag = False
End If
CalculateAllMatrixCorrections = False
Call CalcZAFCalculateExportAll(FormSECONDARY)
If ierror Then Exit Sub
End Sub

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
Call SecondarySave
If ierror Then Exit Sub
Unload FormSECONDARY
End Sub

Private Sub CommandCopyToClipboard_Click()
If Not DebugMode Then On Error Resume Next
' Clipboard (use special function to save graphics methods)
Call SecondaryCopyToClipboard
If ierror Then Exit Sub
End Sub

Private Sub CommandHelp_Click()
If Not DebugMode Then On Error Resume Next
Call IOBrowseHTTP(ProbeSoftwareInternetBrowseMethod%, "https://smf.probesoftware.com/index.php?topic=58.msg214#msg214")
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

Private Sub GridElementList_Click()
If Not DebugMode Then On Error Resume Next
Dim elementrow As Integer
' Determine current element row number
elementrow% = FormSECONDARY.GridElementList.row
' Load k-ratio form
Call CalcZAFSecondaryKratiosLoadForm(elementrow%)
If ierror Then Exit Sub
' Update the element grid
Call CalcZAFSecondaryUpdateList(elementrow%)
If ierror Then Exit Sub
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Not DebugMode Then On Error Resume Next
ImageX1! = x!    ' store for boundary draw
ImageY1! = y!    ' store for boundary draw
Call SecondaryGetBoundary(Int(1), ImageX1!, ImageY1!, ImageX2!, ImageY2!, FormSECONDARY)
If ierror Then Exit Sub
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Not DebugMode Then On Error Resume Next
' Update the stage cursor
Call SecondaryUpdateCursor(x!, y!, FormSECONDARY)
If ierror Then Exit Sub
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Not DebugMode Then On Error Resume Next
ImageX2! = x!
ImageY2! = y!
Call SecondaryGetBoundary(Int(2), ImageX1!, ImageY1!, ImageX2!, ImageY2!, FormSECONDARY)
If ierror Then Exit Sub
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
