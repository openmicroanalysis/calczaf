VERSION 5.00
Begin VB.Form FormADDSTD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Standards to Run"
   ClientHeight    =   4290
   ClientLeft      =   330
   ClientTop       =   1575
   ClientWidth     =   8985
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
   ScaleHeight     =   4290
   ScaleWidth      =   8985
   Begin VB.CommandButton CommandFindNext 
      BackColor       =   &H00C0FFFF&
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3840
      Width           =   2295
   End
   Begin VB.CommandButton CommandHelpAddStd 
      BackColor       =   &H00FF8080&
      Caption         =   "Help"
      Height          =   255
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Click this button to get detailed help from our on-line user forum"
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox TextStandardString 
      Height          =   285
      Left            =   240
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Type a few characters of the standard name and the program will automatically select it"
      Top             =   3480
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "<< Remove Standard from Run"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Remove a single standard from the run"
      Top             =   3840
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Add Standard To Run >>"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Add multiple selected standards to the run"
      Top             =   3360
      Width           =   3015
   End
   Begin VB.ListBox ListAvailableStandards 
      Height          =   2790
      Left            =   120
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Double click to add a single standard to the run"
      Top             =   360
      Width           =   4335
   End
   Begin VB.ListBox ListCurrentStandards 
      Height          =   2790
      Left            =   4560
      Sorted          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Standards in the current run"
      Top             =   360
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7560
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label LabelNumberOfStds 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6360
      TabIndex        =   11
      Top             =   3720
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
      Left            =   6120
      TabIndex        =   10
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Enter Standard To Find:"
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
      TabIndex        =   9
      Top             =   3240
      Width           =   2535
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
' (c) Copyright 1995-2015 by John J. Donovan
Option Explicit

Private Sub Command1_Click()
' Save standards to add to run
If Not DebugMode Then On Error Resume Next
Call AddStdSave
If ierror Then Exit Sub
Unload FormADDSTD
End Sub

Private Sub Command2_Click()
' User clicked cancel in FormADDSTD
If Not DebugMode Then On Error Resume Next
Unload FormADDSTD
Call AddStdCancel   ' reload the original standards
'If ierror Then Exit Sub    ' do not exit on error
icancelload = True
End Sub

Private Sub Command3_Click()
' Add the selected standard to the run
If Not DebugMode Then On Error Resume Next
Call AddStdAdd
If ierror Then Exit Sub
End Sub

Private Sub Command4_Click()
' Remove the selected standard to the run
If Not DebugMode Then On Error Resume Next
Call AddStdRemove
If ierror Then Exit Sub
End Sub

Private Sub CommandFindNext_Click()
If Not DebugMode Then On Error Resume Next
Call StandardFindString(Int(1), FormADDSTD.TextStandardString.Text, FormADDSTD.ListAvailableStandards)
End Sub

Private Sub CommandHelpAddStd_Click()
If Not DebugMode Then On Error Resume Next
Call IOBrowseHTTP(ProbeSoftwareInternetBrowseMethod%, "http://probesoftware.com/smf/index.php?topic=15.0")
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

Private Sub TextStandardString_Change()
If Not DebugMode Then On Error Resume Next
Call StandardFindString(Int(0), FormADDSTD.TextStandardString.Text, FormADDSTD.ListAvailableStandards)
If ierror Then Exit Sub
End Sub
