VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form FormANALYZE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Analyze!"
   ClientHeight    =   5745
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   6225
   Icon            =   "ANALYZE3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSFlexGridLib.MSFlexGrid GridData 
      Height          =   855
      Left            =   3240
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2160
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1508
      _Version        =   393216
      Rows            =   501
      Cols            =   76
   End
   Begin MSFlexGridLib.MSFlexGrid GridStat 
      Height          =   975
      Left            =   120
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2160
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1720
      _Version        =   393216
      Rows            =   8
      Cols            =   76
      ScrollBars      =   1
   End
   Begin VB.TextBox TextDescription 
      BackColor       =   &H8000000B&
      Height          =   615
      Left            =   1680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Sample description information"
      Top             =   1320
      Width           =   4215
   End
   Begin VB.CheckBox CheckOnlyDisplaySamplesWithData 
      Caption         =   "Check1"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Check this box if you do not want to see samples without data listed in the sample list"
      Top             =   360
      Width           =   255
   End
   Begin VB.OptionButton OptionStandard 
      Caption         =   "Standards"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1335
   End
   Begin VB.OptionButton OptionUnknown 
      Caption         =   "Unknowns"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1335
   End
   Begin VB.OptionButton OptionAllSamples 
      Caption         =   "All Samples"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1335
   End
   Begin VB.OptionButton OptionWavescan 
      Caption         =   "Wavescans"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1335
   End
   Begin VB.ListBox ListAnalyze 
      Height          =   645
      Left            =   3360
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   2775
   End
   Begin VB.CheckBox CheckDoNotOutputToLog 
      Caption         =   "Do Not Output To Log"
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
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton CommandNext 
      Height          =   255
      Left            =   1920
      Picture         =   "ANALYZE3.frx":6E7FA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin ComctlLib.StatusBar StatusBarAnal 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5490
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   7329
            TextSave        =   ""
            Key             =   "status"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Cancel"
            TextSave        =   "Cancel"
            Key             =   "cancel"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Next"
            TextSave        =   "Next"
            Key             =   "next"
            Object.Tag             =   ""
         EndProperty
      EndProperty
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
   Begin VB.Label LabelDataType 
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1200
      TabIndex        =   16
      Top             =   3840
      Width           =   4095
   End
   Begin VB.Label LabelZbar 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3960
      TabIndex        =   14
      ToolTipText     =   "Average atomic number"
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label LabelExcess 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1320
      TabIndex        =   13
      ToolTipText     =   "Excess oxygen from calculation (average for oxide calculation)"
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label LabelTotalOxygen 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1320
      TabIndex        =   12
      ToolTipText     =   "Total measured and calculated oxygen (average for oxide calculation)"
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label LabelAtomic 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3960
      TabIndex        =   11
      ToolTipText     =   "Average atomic weight"
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label LabelCalculated 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      ToolTipText     =   "Calculated oxygen only  (average for oxide calculation)"
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label LabelTotal 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3960
      TabIndex        =   9
      ToolTipText     =   "Total weight percent (average)"
      Top             =   4200
      Width           =   855
   End
End
Attribute VB_Name = "FormANALYZE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2017 by John J. Donovan
Option Explicit

Private Sub CommandNext_Click()
If Not DebugMode Then On Error Resume Next
' Nothing to do here
End Sub

Private Sub StatusBarAnal_PanelClick(ByVal Panel As ComctlLib.Panel)
If Not DebugMode Then On Error Resume Next
Select Case Panel.Key
Case "status"
    Exit Sub
Case "cancel"
    Call AnalyzeCancel(FormANALYZE)
    If ierror Then Exit Sub
Case "next"
    Call AnalyzeNext(Int(2))
    If ierror Then Exit Sub
Case Else
    Exit Sub
End Select
End Sub
