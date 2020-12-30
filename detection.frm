VERSION 5.00
Begin VB.Form FormDETECTION 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Model Detection Limits and Counting Times"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows Default
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
      Height          =   615
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Calculation Parameters"
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
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7335
      Begin VB.TextBox TextUnknownBeamCurrent 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Enter the approximate beam current for the unknown sample"
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox TextUnknownBackgroundIntensity 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Enter the approximate background intensity for the unknown sample"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox TextStandardIntensity 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Enter the approximate on-peak intensity for the standard sample"
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox TextStandardWeightPercent 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Enter the weight percent of the emitted element in the standard"
         Top             =   3240
         Width           =   1575
      End
      Begin VB.OLE OLE3 
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         Height          =   495
         Index           =   6
         Left            =   1800
         OleObjectBlob   =   "DETECTION.frx":0000
         TabIndex        =   26
         Top             =   3240
         Width           =   615
      End
      Begin VB.OLE OLE3 
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         Height          =   495
         Index           =   3
         Left            =   1800
         OleObjectBlob   =   "DETECTION.frx":0E18
         TabIndex        =   25
         Top             =   2400
         Width           =   615
      End
      Begin VB.OLE OLE3 
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   1800
         OleObjectBlob   =   "DETECTION.frx":1C30
         TabIndex        =   24
         Top             =   1440
         Width           =   615
      End
      Begin VB.OLE OLE3 
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   1770
         OleObjectBlob   =   "DETECTION.frx":2A48
         TabIndex        =   23
         Top             =   570
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Unknown Beam Current (in nA)"
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
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   6975
      End
      Begin VB.Label Label7 
         Caption         =   "Unknown X-ray Background Intensity (in counts per second per nA)"
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
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   6255
      End
      Begin VB.Label Label3 
         Caption         =   "Standard X-ray Intensity (background corrected in counts per sec per nA)"
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
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   6255
      End
      Begin VB.Label Label1 
         Caption         =   "Weight Percent of Element in Standard (in weight percent)"
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
         Left            =   120
         TabIndex        =   9
         Top             =   3000
         Width           =   6255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Results (assuming 3 sigma statistics)"
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
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   8895
      Begin VB.TextBox TextUnknownOnPeakTime 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Enter the projected count time for the unknown sample"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox TextUnknownWeightPercent 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Enter the predicted concentration of the emitted element in the unknown element in the unknown sample"
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CommandButton CommandPredictCountTime 
         BackColor       =   &H0080FFFF&
         Caption         =   "Predict Count Time To Detect a Specified Concentration"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   3120
         Width           =   5535
      End
      Begin VB.CommandButton CommandPredictDetectionLimit 
         BackColor       =   &H0080FFFF&
         Caption         =   "Predict Detection Limit For a Specified Integration Time"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1200
         Width           =   5535
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "yields"
         Height          =   255
         Left            =   2520
         TabIndex        =   30
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "yields"
         Height          =   255
         Left            =   2520
         TabIndex        =   29
         Top             =   720
         Width           =   615
      End
      Begin VB.OLE OLE3 
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         Height          =   495
         Index           =   8
         Left            =   1800
         OleObjectBlob   =   "DETECTION.frx":3860
         TabIndex        =   28
         Top             =   2520
         Width           =   615
      End
      Begin VB.OLE OLE3 
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         Height          =   495
         Index           =   7
         Left            =   1800
         OleObjectBlob   =   "DETECTION.frx":4678
         TabIndex        =   27
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "seconds"
         Height          =   255
         Left            =   4560
         TabIndex        =   22
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "weight percent"
         Height          =   255
         Left            =   4560
         TabIndex        =   21
         Top             =   720
         Width           =   1095
      End
      Begin VB.OLE OLE2 
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         Height          =   1095
         Left            =   5760
         OleObjectBlob   =   "DETECTION.frx":5490
         TabIndex        =   20
         Top             =   840
         Width           =   3015
      End
      Begin VB.OLE OLE1 
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         Height          =   975
         Left            =   5760
         OleObjectBlob   =   "DETECTION.frx":66A8
         TabIndex        =   19
         Top             =   2760
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "Unknown Peak (or Background) Integration Time (in seconds)"
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
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   8655
      End
      Begin VB.Label Label6 
         Caption         =   "Concentration of Element at 99% confidence (in weight %)"
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
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   6015
      End
      Begin VB.Label LabelTimePredicted 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   14
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label LabelWeightPercentDetected 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Contributed by John Fournelle"
      Height          =   375
      Left            =   7560
      TabIndex        =   31
      Top             =   3600
      Width           =   1455
   End
End
Attribute VB_Name = "FormDETECTION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2021 by John J. Donovan
Option Explicit

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
Call DetectionSave
If ierror Then Exit Sub
Unload FormDETECTION
End Sub

Private Sub CommandPredictCountTime_Click()
If Not DebugMode Then On Error Resume Next
Call DetectionSave
If ierror Then Exit Sub
Call DetectionCalculateCountTime
If ierror Then Exit Sub
Call DetectionPrint(Int(2))
If ierror Then Exit Sub
End Sub

Private Sub CommandPredictDetectionLimit_Click()
If Not DebugMode Then On Error Resume Next
Call DetectionSave
If ierror Then Exit Sub
Call DetectionCalculateConcentration
If ierror Then Exit Sub
Call DetectionPrint(Int(1))
If ierror Then Exit Sub
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormDETECTION)
HelpContextID = IOGetHelpContextID("FormDETECTION")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub TextStandardIntensity_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextStandardWeightPercent_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextUnknownBackgroundIntensity_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextUnknownBeamCurrent_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextUnknownOnPeakTime_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextUnknownWeightPercent_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub
