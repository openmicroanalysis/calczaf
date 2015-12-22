VERSION 5.00
Begin VB.Form FormEMP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Empirical"
   ClientHeight    =   6090
   ClientLeft      =   1140
   ClientTop       =   2010
   ClientWidth     =   14880
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6090
   ScaleWidth      =   14880
   Begin VB.TextBox TextReNormalizeStandard 
      Height          =   285
      Left            =   12480
      TabIndex        =   11
      ToolTipText     =   $"EMP.frx":0000
      Top             =   3480
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox TextReNormalizeFactor 
      Height          =   285
      Left            =   10080
      TabIndex        =   9
      ToolTipText     =   "Enter the APF of the element in the primary standard to re-normalize the factors to the standard"
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000C000&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Delete From Run"
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Remove selected empirical value from run"
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Add To Run >>"
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Add selected empirical value to run"
      Top             =   3240
      Width           =   2055
   End
   Begin VB.ListBox ListCurrentEmp 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   7200
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   $"EMP.frx":00DE
      Top             =   720
      Width           =   7575
   End
   Begin VB.ListBox ListAvailableEmp 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Available empirical MAC or APF values (edit ASCII file to add additional values)"
      Top             =   720
      Width           =   6735
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   13560
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label LabelMAC 
      Alignment       =   2  'Center
      Caption         =   $"EMP.frx":016C
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   15
      Top             =   3240
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.OLE OLE2 
      BackStyle       =   0  'Transparent
      Class           =   "Word.Document.8"
      Enabled         =   0   'False
      Height          =   975
      Left            =   6360
      OleObjectBlob   =   "EMP.frx":02CA
      TabIndex        =   14
      Top             =   4440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label LabelAPF 
      Alignment       =   2  'Center
      Caption         =   $"EMP.frx":A0E2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   13
      Top             =   3240
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.Label LabelReNormalizeStandard 
      Alignment       =   2  'Center
      Caption         =   "Re-Normalization Standard"
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
      Left            =   12480
      TabIndex        =   12
      Top             =   3240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label LabelReNormalize 
      Alignment       =   2  'Center
      Caption         =   $"EMP.frx":A519
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9240
      TabIndex        =   10
      Top             =   3840
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label LabelReNormalizeFactor 
      Alignment       =   2  'Center
      Caption         =   "Re-Normalization Factor"
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
      Left            =   9600
      TabIndex        =   8
      Top             =   3240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label LabelCurrent 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Current Empirical"
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
      Height          =   615
      Left            =   7200
      TabIndex        =   4
      Top             =   120
      Width           =   7455
   End
   Begin VB.Label LabelAvailable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Available Empirical"
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
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "FormEMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2015 by John J. Donovan
Option Explicit

Private Sub Command1_Click()
If Not DebugMode Then On Error Resume Next
Unload FormEMP
icancelload = True
End Sub

Private Sub Command2_Click()
' OK click
If Not DebugMode Then On Error Resume Next
Call EmpSave
If ierror Then Exit Sub
Unload FormEMP
End Sub

Private Sub Command3_Click()
' Load the selected value
If Not DebugMode Then On Error Resume Next
Call EmpAddEmp
If ierror Then Exit Sub
End Sub

Private Sub Command4_Click()
' Delete a specified empirical AMC/APF
If Not DebugMode Then On Error Resume Next
Call EmpDeleteEmp
If ierror Then Exit Sub
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
icancelload = False
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormEMP)
If EmpTypeFlag% = 1 Then
HelpContextID = IOGetHelpContextID("FormEMPMAC")
Else
HelpContextID = IOGetHelpContextID("FormEMPAPF")
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub ListAvailableEmp_DblClick()
' Load the selected value
If Not DebugMode Then On Error Resume Next
Call EmpAddEmp
If ierror Then Exit Sub
End Sub

Private Sub ListCurrentEmp_Click()
Call EmpLoadReNormalization
If ierror Then Exit Sub
End Sub

Private Sub TextReNormalizeFactor_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextReNormalizeStandard_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub
