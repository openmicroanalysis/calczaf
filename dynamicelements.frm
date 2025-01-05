VERSION 5.00
Begin VB.Form FormDynamicElements 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calculate Specified Elements Dynamically Based on K-Ratio Limits"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   15615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox ComboOxygenByStoichiometryOperator2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12120
      Style           =   2  'Dropdown List
      TabIndex        =   78
      Top             =   4200
      Width           =   735
   End
   Begin VB.ComboBox ComboOxygenByStoichiometryOperator1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8520
      Style           =   2  'Dropdown List
      TabIndex        =   77
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox TextOxygenByStoichiometryValue 
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
      Index           =   2
      Left            =   14640
      TabIndex        =   76
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox TextOxygenByStoichiometryValue 
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
      Index           =   1
      Left            =   11040
      TabIndex        =   75
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox TextOxygenByStoichiometryValue 
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
      Index           =   0
      Left            =   7440
      TabIndex        =   74
      Top             =   4200
      Width           =   855
   End
   Begin VB.ComboBox ComboOxygenByStoichiometryGreaterLess 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   13920
      Style           =   2  'Dropdown List
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   4200
      Width           =   615
   End
   Begin VB.ComboBox ComboOxygenByStoichiometryGreaterLess 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   10320
      Style           =   2  'Dropdown List
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   4200
      Width           =   615
   End
   Begin VB.ComboBox ComboOxygenByStoichiometryGreaterLess 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   6720
      Style           =   2  'Dropdown List
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   4200
      Width           =   615
   End
   Begin VB.ComboBox ComboOxygenByStoichiometryElement 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   13080
      Style           =   2  'Dropdown List
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   4200
      Width           =   735
   End
   Begin VB.ComboBox ComboOxygenByStoichiometryElement 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   9480
      Style           =   2  'Dropdown List
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   4200
      Width           =   735
   End
   Begin VB.ComboBox ComboOxygenByStoichiometryElement 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   5880
      Style           =   2  'Dropdown List
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   4200
      Width           =   735
   End
   Begin VB.CheckBox CheckOxygenByStoichiometry 
      Caption         =   "Dynamically Calculate Oxygen by Stoichiometry"
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
      Left            =   240
      TabIndex        =   67
      TabStop         =   0   'False
      ToolTipText     =   "Calculate oxygen by stoichiometry dynamically based on the specified k-ratio criteria"
      Top             =   4200
      Width           =   4815
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
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   66
      ToolTipText     =   "Click this button to get detailed help from our on-line user forum"
      Top             =   480
      Width           =   1215
   End
   Begin VB.ComboBox ComboExcessOxygenByDroopOperator2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12120
      Style           =   2  'Dropdown List
      TabIndex        =   65
      Top             =   3600
      Width           =   735
   End
   Begin VB.ComboBox ComboExcessOxygenByDroopOperator1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8520
      Style           =   2  'Dropdown List
      TabIndex        =   64
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox TextExcessOxygenByDroopValue 
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
      Index           =   2
      Left            =   14640
      TabIndex        =   63
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox TextExcessOxygenByDroopValue 
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
      Index           =   1
      Left            =   11040
      TabIndex        =   62
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox TextExcessOxygenByDroopValue 
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
      Index           =   0
      Left            =   7440
      TabIndex        =   61
      Top             =   3600
      Width           =   855
   End
   Begin VB.ComboBox ComboExcessOxygenByDroopGreaterLess 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   13920
      Style           =   2  'Dropdown List
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   3600
      Width           =   615
   End
   Begin VB.ComboBox ComboExcessOxygenByDroopGreaterLess 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   10320
      Style           =   2  'Dropdown List
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   3600
      Width           =   615
   End
   Begin VB.ComboBox ComboExcessOxygenByDroopGreaterLess 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   6720
      Style           =   2  'Dropdown List
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   3600
      Width           =   615
   End
   Begin VB.ComboBox ComboExcessOxygenByDroopElement 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   13080
      Style           =   2  'Dropdown List
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   3600
      Width           =   735
   End
   Begin VB.ComboBox ComboExcessOxygenByDroopElement 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   9480
      Style           =   2  'Dropdown List
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   3600
      Width           =   735
   End
   Begin VB.ComboBox ComboExcessOxygenByDroopElement 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   5880
      Style           =   2  'Dropdown List
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   3600
      Width           =   735
   End
   Begin VB.CheckBox CheckExcessOxygenByDroop 
      Caption         =   "Dynamically Calculate Excess Oxygen from Ferric/Ferrous Ratio"
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
      Left            =   240
      TabIndex        =   54
      ToolTipText     =   "Calculate excess oxygen by charge balance (Droop, 1987) dynamically based on the specified k-ratio criteria"
      Top             =   3480
      Width           =   5175
   End
   Begin VB.ComboBox ComboRelativeOperator2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12120
      Style           =   2  'Dropdown List
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   3120
      Width           =   735
   End
   Begin VB.ComboBox ComboRelativeOperator1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8520
      Style           =   2  'Dropdown List
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   3120
      Width           =   735
   End
   Begin VB.ComboBox ComboStoichiometryOperator2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12120
      Style           =   2  'Dropdown List
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   2640
      Width           =   735
   End
   Begin VB.ComboBox ComboStoichiometryOperator1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8520
      Style           =   2  'Dropdown List
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   2640
      Width           =   735
   End
   Begin VB.ComboBox ComboDifferenceFormulaOperator2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12120
      Style           =   2  'Dropdown List
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   2160
      Width           =   735
   End
   Begin VB.ComboBox ComboDifferenceFormulaOperator1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8520
      Style           =   2  'Dropdown List
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox TextRelativeValue 
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
      Index           =   2
      Left            =   14640
      TabIndex        =   44
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox TextRelativeValue 
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
      Index           =   1
      Left            =   11040
      TabIndex        =   43
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox TextRelativeValue 
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
      Index           =   0
      Left            =   7440
      TabIndex        =   42
      Top             =   3120
      Width           =   855
   End
   Begin VB.ComboBox ComboRelativeGreaterLess 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   13920
      Style           =   2  'Dropdown List
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   3120
      Width           =   615
   End
   Begin VB.ComboBox ComboRelativeGreaterLess 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   10320
      Style           =   2  'Dropdown List
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   3120
      Width           =   615
   End
   Begin VB.ComboBox ComboRelativeGreaterLess 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   6720
      Style           =   2  'Dropdown List
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox TextStoichiometryValue 
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
      Index           =   2
      Left            =   14640
      TabIndex        =   38
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox TextStoichiometryValue 
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
      Index           =   1
      Left            =   11040
      TabIndex        =   37
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox TextStoichiometryValue 
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
      Index           =   0
      Left            =   7440
      TabIndex        =   36
      Top             =   2640
      Width           =   855
   End
   Begin VB.ComboBox ComboStoichiometryGreaterLess 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   13920
      Style           =   2  'Dropdown List
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   2640
      Width           =   615
   End
   Begin VB.ComboBox ComboStoichiometryGreaterLess 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   10320
      Style           =   2  'Dropdown List
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2640
      Width           =   615
   End
   Begin VB.ComboBox ComboStoichiometryGreaterLess 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   6720
      Style           =   2  'Dropdown List
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox TextDifferenceFormulaValue 
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
      Index           =   2
      Left            =   14640
      TabIndex        =   32
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox TextDifferenceFormulaValue 
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
      Index           =   1
      Left            =   11040
      TabIndex        =   31
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox TextDifferenceFormulaValue 
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
      Index           =   0
      Left            =   7440
      TabIndex        =   30
      Top             =   2160
      Width           =   855
   End
   Begin VB.ComboBox ComboDifferenceFormulaGreaterLess 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   13920
      Style           =   2  'Dropdown List
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2160
      Width           =   615
   End
   Begin VB.ComboBox ComboDifferenceFormulaGreaterLess 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   10320
      Style           =   2  'Dropdown List
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2160
      Width           =   615
   End
   Begin VB.ComboBox ComboDifferenceFormulaGreaterLess 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   6720
      Style           =   2  'Dropdown List
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2160
      Width           =   615
   End
   Begin VB.ComboBox ComboRelativeElement 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   13080
      Style           =   2  'Dropdown List
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   3120
      Width           =   735
   End
   Begin VB.ComboBox ComboRelativeElement 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   9480
      Style           =   2  'Dropdown List
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3120
      Width           =   735
   End
   Begin VB.ComboBox ComboRelativeElement 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   5880
      Style           =   2  'Dropdown List
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3120
      Width           =   735
   End
   Begin VB.ComboBox ComboStoichiometryElement 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   13080
      Style           =   2  'Dropdown List
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2640
      Width           =   735
   End
   Begin VB.ComboBox ComboStoichiometryElement 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   9480
      Style           =   2  'Dropdown List
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2640
      Width           =   735
   End
   Begin VB.ComboBox ComboStoichiometryElement 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   5880
      Style           =   2  'Dropdown List
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2640
      Width           =   735
   End
   Begin VB.ComboBox ComboDifferenceFormulaElement 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   13080
      Style           =   2  'Dropdown List
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2160
      Width           =   735
   End
   Begin VB.ComboBox ComboDifferenceFormulaElement 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   9480
      Style           =   2  'Dropdown List
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2160
      Width           =   735
   End
   Begin VB.ComboBox ComboDifferenceFormulaElement 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   5880
      Style           =   2  'Dropdown List
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox TextDifferenceValue 
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
      Index           =   2
      Left            =   14640
      TabIndex        =   17
      Top             =   1680
      Width           =   855
   End
   Begin VB.ComboBox ComboDifferenceOperator2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12120
      Style           =   2  'Dropdown List
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1680
      Width           =   735
   End
   Begin VB.ComboBox ComboDifferenceOperator1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8520
      Style           =   2  'Dropdown List
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox TextDifferenceValue 
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
      Index           =   1
      Left            =   11040
      TabIndex        =   14
      Top             =   1680
      Width           =   855
   End
   Begin VB.ComboBox ComboDifferenceGreaterLess 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   13920
      Style           =   2  'Dropdown List
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1680
      Width           =   615
   End
   Begin VB.ComboBox ComboDifferenceGreaterLess 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   10320
      Style           =   2  'Dropdown List
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1680
      Width           =   615
   End
   Begin VB.ComboBox ComboDifferenceGreaterLess 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   6720
      Style           =   2  'Dropdown List
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox TextDifferenceValue 
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
      Index           =   0
      Left            =   7440
      TabIndex        =   10
      Top             =   1680
      Width           =   855
   End
   Begin VB.ComboBox ComboDifferenceElement 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   13080
      Style           =   2  'Dropdown List
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1680
      Width           =   735
   End
   Begin VB.ComboBox ComboDifferenceElement 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   9480
      Style           =   2  'Dropdown List
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1680
      Width           =   735
   End
   Begin VB.ComboBox ComboDifferenceElement 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   5880
      Style           =   2  'Dropdown List
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1680
      Width           =   735
   End
   Begin VB.CheckBox CheckRelative 
      Caption         =   "Dynamically Calculate Stoichiometry To Another Element"
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
      Left            =   240
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Calculate an element by stoichiometry to another element dynamically based on the specified k-ratio criteria"
      Top             =   3120
      Width           =   5175
   End
   Begin VB.CheckBox CheckStoichiometry 
      Caption         =   "Dynamically Calculate Stoichiometry To Oxygen"
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
      Left            =   240
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Calculate an element by stoichiometry to stoichiometric oxygen dynamically based on the specified k-ratio criteria"
      Top             =   2640
      Width           =   4815
   End
   Begin VB.CheckBox CheckDifferenceFormula 
      Caption         =   "Dynamically Calculate Formula By Difference"
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
      Left            =   240
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Calculate formula by difference dynamically based on the specified k-ratio criteria"
      Top             =   2160
      Width           =   4815
   End
   Begin VB.CheckBox CheckDifference 
      Caption         =   "Dynamically Calculate Element By Difference"
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
      Left            =   240
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Calculate an element by difference from 100% dynamically based on the specified k-ratio criteria"
      Top             =   1680
      Width           =   4815
   End
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
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
      Left            =   12720
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton CommandOK 
      BackColor       =   &H00C0FFC0&
      Caption         =   "OK"
      Default         =   -1  'True
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
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label LabelKRatio 
      Alignment       =   2  'Center
      Caption         =   "K-Ratio"
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
      Left            =   14640
      TabIndex        =   53
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label LabelKRatio 
      Alignment       =   2  'Center
      Caption         =   "K-Ratio"
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
      Left            =   11040
      TabIndex        =   52
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label LabelKRatio 
      Alignment       =   2  'Center
      Caption         =   "K-Ratio"
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
      Left            =   7440
      TabIndex        =   51
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label LabelInstructions 
      Alignment       =   2  'Center
      Height          =   1335
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   11295
   End
End
Attribute VB_Name = "FormDynamicElements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2025 by John J. Donovan
Option Explicit

Private Sub CommandCancel_Click()
If Not DebugMode Then On Error Resume Next
Unload FormDynamicElements
End Sub

Private Sub CommandHelp_Click()
If Not DebugMode Then On Error Resume Next
Call IOBrowseHTTP(ProbeSoftwareInternetBrowseMethod%, "https://smf.probesoftware.com/index.php?topic=1647.0")
If ierror Then Exit Sub
End Sub

Private Sub CommandOK_Click()
If Not DebugMode Then On Error Resume Next
Call DynamicElementsSave
If ierror Then Exit Sub
Unload FormDynamicElements
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
icancelload = False
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormDynamicElements)
HelpContextID = IOGetHelpContextID("FormDynamicElements")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub TextDifferenceFormulaValue_GotFocus(Index As Integer)
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextDifferenceValue_GotFocus(Index As Integer)
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextExcessOxygenByDroopValue_GotFocus(Index As Integer)
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextOxygenByStoichiometryValue_GotFocus(Index As Integer)
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextRelativeValue_GotFocus(Index As Integer)
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextStoichiometryValue_GotFocus(Index As Integer)
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub
