VERSION 5.00
Begin VB.Form FormEXCELOPTIONS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Excel Spreadsheet Output Options"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   720
      Width           =   1095
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Excel Output Options"
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
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.OptionButton OptionExcelOutputOption 
         Caption         =   "Output Normalized Oxide Percents"
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
         Index           =   6
         Left            =   240
         TabIndex        =   9
         ToolTipText     =   "Select Excel output option"
         Top             =   2040
         Width           =   3975
      End
      Begin VB.OptionButton OptionExcelOutputOption 
         Caption         =   "Output Normalized Elemental Percents"
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
         Index           =   5
         Left            =   240
         TabIndex        =   8
         ToolTipText     =   "Select Excel output option"
         Top             =   1800
         Width           =   3975
      End
      Begin VB.OptionButton OptionExcelOutputOption 
         Caption         =   "Output Elemental Weight Percents"
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
         TabIndex        =   5
         ToolTipText     =   "Select Excel output option"
         Top             =   360
         Width           =   3615
      End
      Begin VB.OptionButton OptionExcelOutputOption 
         Caption         =   "Output Elemental K-Ratios"
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
         TabIndex        =   4
         ToolTipText     =   "Select Excel output option"
         Top             =   600
         Width           =   3735
      End
      Begin VB.OptionButton OptionExcelOutputOption 
         Caption         =   "Output Atomic Percents"
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
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "Select Excel output option"
         Top             =   960
         Width           =   3615
      End
      Begin VB.OptionButton OptionExcelOutputOption 
         Caption         =   "Output Oxide Weight Percents"
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
         Left            =   240
         TabIndex        =   2
         ToolTipText     =   "Select Excel output option"
         Top             =   1200
         Width           =   3735
      End
      Begin VB.OptionButton OptionExcelOutputOption 
         Caption         =   "Output Formula Coefficients"
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
         Index           =   4
         Left            =   240
         TabIndex        =   1
         ToolTipText     =   "Select Excel output option"
         Top             =   1440
         Width           =   3975
      End
   End
End
Attribute VB_Name = "FormEXCELOPTIONS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2016 by John J. Donovan

Option Explicit

Private Sub CommandCancel_Click()
If Not DebugMode Then On Error Resume Next
Unload FormEXCELOPTIONS
End Sub

Private Sub CommandOK_Click()
If Not DebugMode Then On Error Resume Next
Call CalcZAFExcelOptionsSave
If ierror Then Exit Sub
Unload FormEXCELOPTIONS
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
icancelload = False
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormEXCELOPTIONS)
HelpContextID = IOGetHelpContextID("FormEXCELOPTIONS")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub
