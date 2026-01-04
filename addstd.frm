VERSION 5.00
Begin VB.Form FormADDSTD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Standards to Run"
   ClientHeight    =   7080
   ClientLeft      =   330
   ClientTop       =   1575
   ClientWidth     =   10440
   ControlBox      =   0   'False
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7080
   ScaleWidth      =   10440
   Begin VB.CheckBox CheckMaterialType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   29
      Left            =   8400
      TabIndex        =   49
      ToolTipText     =   "Check this material type to filter the available standard list"
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CheckBox CheckMaterialType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   28
      Left            =   6360
      TabIndex        =   48
      ToolTipText     =   "Check this material type to filter the available standard list"
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CheckBox CheckMaterialType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   27
      Left            =   4320
      TabIndex        =   47
      ToolTipText     =   "Check this material type to filter the available standard list"
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CheckBox CheckMaterialType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   26
      Left            =   2280
      TabIndex        =   46
      ToolTipText     =   "Check this material type to filter the available standard list"
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CheckBox CheckMaterialType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   25
      Left            =   240
      TabIndex        =   45
      ToolTipText     =   "Check this material type to filter the available standard list"
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CheckBox CheckMaterialType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   24
      Left            =   8400
      TabIndex        =   44
      ToolTipText     =   "Check this material type to filter the available standard list"
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CheckBox CheckMaterialType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   6360
      TabIndex        =   43
      ToolTipText     =   "Check this material type to filter the available standard list"
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CheckBox CheckMaterialType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   22
      Left            =   4320
      TabIndex        =   42
      ToolTipText     =   "Check this material type to filter the available standard list"
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CheckBox CheckMaterialType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   21
      Left            =   2280
      TabIndex        =   41
      ToolTipText     =   "Check this material type to filter the available standard list"
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CheckBox CheckMaterialType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   240
      TabIndex        =   40
      ToolTipText     =   "Check this material type to filter the available standard list"
      Top             =   5880
      Width           =   1815
   End
   Begin VB.TextBox TextMountNames 
      Height          =   285
      Left            =   6840
      TabIndex        =   38
      ToolTipText     =   "Enter a few characters in the standard mount name to filter the available standard list"
      Top             =   6720
      Width           =   3375
   End
   Begin VB.CommandButton CommandLoadFromPOS 
      Caption         =   "Load POS"
      Height          =   375
      Left            =   7680
      TabIndex        =   37
      ToolTipText     =   "Load standards from a selected .POS file"
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CheckBox CheckMaterialType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   19
      Left            =   8400
      TabIndex        =   27
      ToolTipText     =   "Check this material type to filter the available standard list"
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CheckBox CheckMaterialType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   18
      Left            =   6360
      TabIndex        =   28
      ToolTipText     =   "Check this material type to filter the available standard list"
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CheckBox CheckMaterialType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   4320
      TabIndex        =   29
      ToolTipText     =   "Check this material type to filter the available standard list"
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CheckBox CheckMaterialType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   2280
      TabIndex        =   30
      ToolTipText     =   "Check this material type to filter the available standard list"
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CheckBox CheckMaterialType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   240
      TabIndex        =   31
      ToolTipText     =   "Check this material type to filter the available standard list"
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CheckBox CheckMaterialType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   8400
      TabIndex        =   32
      ToolTipText     =   "Check this material type to filter the available standard list"
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CheckBox CheckMaterialType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   6360
      TabIndex        =   33
      ToolTipText     =   "Check this material type to filter the available standard list"
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CheckBox CheckMaterialType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   4320
      TabIndex        =   34
      ToolTipText     =   "Check this material type to filter the available standard list"
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CheckBox CheckMaterialType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   2280
      TabIndex        =   35
      ToolTipText     =   "Check this material type to filter the available standard list"
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CheckBox CheckMaterialType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   36
      ToolTipText     =   "Check this material type to filter the available standard list"
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CheckBox CheckMaterialType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   8400
      TabIndex        =   26
      ToolTipText     =   "Check this material type to filter the available standard list"
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CheckBox CheckMaterialType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   6360
      TabIndex        =   25
      ToolTipText     =   "Check this material type to filter the available standard list"
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CheckBox CheckMaterialType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   4320
      TabIndex        =   24
      ToolTipText     =   "Check this material type to filter the available standard list"
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CheckBox CheckMaterialType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   2280
      TabIndex        =   23
      ToolTipText     =   "Check this material type to filter the available standard list"
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CheckBox CheckMaterialType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   22
      ToolTipText     =   "Check this material type to filter the available standard list"
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CheckBox CheckMaterialType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   8400
      TabIndex        =   21
      ToolTipText     =   "Check this material type to filter the available standard list"
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CheckBox CheckMaterialType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   6360
      TabIndex        =   20
      ToolTipText     =   "Check this material type to filter the available standard list"
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CheckBox CheckMaterialType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   19
      ToolTipText     =   "Check this material type to filter the available standard list"
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CheckBox CheckMaterialType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   18
      ToolTipText     =   "Check this material type to filter the available standard list"
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CheckBox CheckMaterialType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   17
      ToolTipText     =   "Check this material type to filter the available standard list"
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox TextStandardNumber 
      Height          =   285
      Left            =   2040
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Type a few characters of the standard name and the program will automatically select it"
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton CommandFindNextNumber 
      BackColor       =   &H0080FFFF&
      Caption         =   "Next Match"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton CommandFindNextString 
      BackColor       =   &H0080FFFF&
      Caption         =   "Next Match"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton CommandHelpAddStd 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Help"
      Height          =   375
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Click this button to get detailed help from our on-line user forum"
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox TextStandardString 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Type a few characters of the standard name and the program will automatically select it"
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton CommandRemoveStandardFromRun 
      Caption         =   "<< Remove Standard from Run"
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Remove a single standard from the run"
      Top             =   3840
      Width           =   2775
   End
   Begin VB.CommandButton CommandAddStandardToRun 
      BackColor       =   &H0080FFFF&
      Caption         =   "Add Standard To Run >>"
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Add multiple selected standards to the run"
      Top             =   3240
      Width           =   2775
   End
   Begin VB.ListBox ListAvailableStandards 
      Height          =   2790
      Left            =   120
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Double click to add a single standard to the run (single click to only output composition to log window)"
      Top             =   360
      Width           =   5055
   End
   Begin VB.ListBox ListCurrentStandards 
      Height          =   2790
      Left            =   5280
      Sorted          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Standards in the current run"
      Top             =   360
      Width           =   5055
   End
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   9000
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton CommandOK 
      BackColor       =   &H00C0FFC0&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Enter name of standard mount to list only those standards in that standard mount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   39
      Top             =   6720
      Width           =   6135
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   10320
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   10320
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label5 
      Caption         =   "Enter Number To Find:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   16
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label LabelNumberOfStds 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7320
      TabIndex        =   11
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Number Of Stds"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   10
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Enter String To Find:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Available Standards in Database (multi-select)"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Current Standards in Run"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "FormADDSTD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2026 by John J. Donovan
Option Explicit

Private Sub CheckMaterialType_Click(Index As Integer)
If Not DebugMode Then On Error Resume Next
Call AddStdMaterialTypeFilter
If ierror Then Exit Sub
End Sub

Private Sub CommandAddStandardToRun_Click()
' Add the selected standard to the run
If Not DebugMode Then On Error Resume Next
Call AddStdAdd
If ierror Then Exit Sub
End Sub

Private Sub CommandCancel_Click()
If Not DebugMode Then On Error Resume Next
Unload FormADDSTD
Call AddStdCancel   ' reload the original standards
'If ierror Then Exit Sub    ' do not exit on error
icancelload = True
End Sub

Private Sub CommandFindNextNumber_Click()
If Not DebugMode Then On Error Resume Next
Call StandardFindNumber(Int(1), FormADDSTD.TextStandardNumber.Text, FormADDSTD.ListAvailableStandards)
If ierror Then Exit Sub
End Sub

Private Sub CommandFindNextString_Click()
If Not DebugMode Then On Error Resume Next
Call StandardFindString(Int(1), FormADDSTD.TextStandardString.Text, FormADDSTD.ListAvailableStandards)
If ierror Then Exit Sub
End Sub

Private Sub CommandHelpAddStd_Click()
If Not DebugMode Then On Error Resume Next
Call IOBrowseHTTP(ProbeSoftwareInternetBrowseMethod%, "https://smf.probesoftware.com/index.php?topic=15.0")
If ierror Then Exit Sub
End Sub

Private Sub CommandLoadFromPOS_Click()
If Not DebugMode Then On Error Resume Next
Call AddStdImportPOS
If ierror Then Exit Sub
End Sub

Private Sub CommandOK_Click()
If Not DebugMode Then On Error Resume Next
Call AddStdSave
If ierror Then Exit Sub
Unload FormADDSTD
End Sub

Private Sub CommandRemoveStandardFromRun_Click()
' Remove the selected standard to the run
If Not DebugMode Then On Error Resume Next
Call AddStdRemove
If ierror Then Exit Sub
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
icancelload = False
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormADDSTD)
HelpContextID = IOGetHelpContextID("FormADDSTD")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub ListAvailableStandards_Click()
If Not DebugMode Then On Error Resume Next
Dim stdnum As Integer

' Get standard from listbox
If FormADDSTD.ListAvailableStandards.ListIndex < 0 Then Exit Sub
stdnum% = FormADDSTD.ListAvailableStandards.ItemData(FormADDSTD.ListAvailableStandards.ListIndex)

' Display standard data
If stdnum% > 0 Then Call StandardTypeStandard(stdnum%)
If ierror Then Exit Sub
End Sub

Private Sub ListAvailableStandards_DblClick()
' Add the selected standard to the run
If Not DebugMode Then On Error Resume Next
Call AddStdAdd
If ierror Then Exit Sub
End Sub

Private Sub ListCurrentStandards_DblClick()
' Remove the selected standard to the run
If Not DebugMode Then On Error Resume Next
Call AddStdRemove
If ierror Then Exit Sub
End Sub

Private Sub TextMountNames_Change()
If Not DebugMode Then On Error Resume Next
Call AddStdMountNamesFilter
If ierror Then Exit Sub
End Sub

Private Sub TextStandardNumber_Change()
If Not DebugMode Then On Error Resume Next
Call StandardFindNumber(Int(0), FormADDSTD.TextStandardNumber.Text, FormADDSTD.ListAvailableStandards)
If ierror Then Exit Sub
End Sub

Private Sub TextStandardString_Change()
If Not DebugMode Then On Error Resume Next
Call StandardFindString(Int(0), FormADDSTD.TextStandardString.Text, FormADDSTD.ListAvailableStandards)
If ierror Then Exit Sub
End Sub
