VERSION 5.00
Begin VB.Form FormSTDCOMP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Standard Entry"
   ClientHeight    =   3570
   ClientLeft      =   750
   ClientTop       =   4290
   ClientWidth     =   6600
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
   ScaleHeight     =   3570
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton CommandOK 
      BackColor       =   &H00C0FFC0&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Standard Composition for Calculation"
      ClipControls    =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5055
      Begin VB.ListBox ListAvailableStandards 
         Height          =   2790
         Left            =   120
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Double click to add a single standard to the run"
         Top             =   360
         Width           =   4815
      End
   End
End
Attribute VB_Name = "FormSTDCOMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2021 by John J. Donovan
Option Explicit

Private Sub CommandCancel_Click()
If Not DebugMode Then On Error Resume Next
Unload FormSTDCOMP
ierror = True
End Sub

Private Sub CommandOK_Click()
If Not DebugMode Then On Error Resume Next
Call FormulaSaveStdComp
If ierror Then Exit Sub
End Sub

Private Sub Form_Activate()
If Not DebugMode Then On Error Resume Next
' Get available standard names and numbers from database
Call StandardGetMDBIndex
If ierror Then Exit Sub
' List the available standards
Call StandardLoadList(FormSTDCOMP.ListAvailableStandards)
If ierror Then Exit Sub
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call MiscCenterForm(FormSTDCOMP)
Call MiscLoadIcon(FormSTDCOMP)
HelpContextID = IOGetHelpContextID("FormSTDCOMP")
End Sub

Private Sub ListAvailableStandards_DblClick()
If Not DebugMode Then On Error Resume Next
Dim stdnum As Integer

' Get standard from listbox
If FormSTDCOMP.ListAvailableStandards.ListIndex < 0 Then Exit Sub
stdnum% = FormSTDCOMP.ListAvailableStandards.ItemData(FormSTDCOMP.ListAvailableStandards.ListIndex)

' Display standard data
If stdnum% > 0 Then Call StandardTypeStandard(stdnum%)
If ierror Then Exit Sub
End Sub
