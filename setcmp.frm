VERSION 5.00
Begin VB.Form FormSETCMP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Element Properties"
   ClientHeight    =   2130
   ClientLeft      =   1410
   ClientTop       =   1770
   ClientWidth     =   7305
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
   ScaleHeight     =   2130
   ScaleWidth      =   7305
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
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
      Left            =   6120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enter Element Properties and Weight Percent For"
      ClipControls    =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5895
      Begin VB.ComboBox ComboCrystal 
         Height          =   315
         Left            =   3000
         TabIndex        =   16
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox TextCharge 
         Height          =   285
         Left            =   4440
         TabIndex        =   5
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox TextComposition 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   2775
      End
      Begin VB.ComboBox ComboOxygens 
         Height          =   315
         Left            =   4440
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox ComboCations 
         Height          =   315
         Left            =   3000
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox ComboXRay 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox ComboElement 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Crystal (default)"
         Height          =   255
         Left            =   2880
         TabIndex        =   17
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Charge"
         Height          =   255
         Left            =   4440
         TabIndex        =   15
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label LabelComposition 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Enter Composition In "
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Oxygens"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4440
         TabIndex        =   11
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Cations"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3000
         TabIndex        =   12
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "X-Ray (default)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Element"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FormSETCMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2016 by John J. Donovan
Option Explicit

Private Sub ComboElement_Change()
If Not DebugMode Then On Error Resume Next
' Update
Call GetCmpSetCmpUpdateCombo
If ierror Then Exit Sub
End Sub

Private Sub ComboElement_Click()
If Not DebugMode Then On Error Resume Next
' Update
Call GetCmpSetCmpUpdateCombo
If ierror Then Exit Sub
End Sub

Private Sub Command1_Click()
If Not DebugMode Then On Error Resume Next

' User clicked OK
Call GetCmpSetCmpSave
If ierror Then Exit Sub

Unload FormSETCMP
DoEvents

' Remove blank rows
Call GetCmpSave
If ierror Then Exit Sub

' Reload the entire grid in case an element was deleted
Call GetCmpLoadGrid
If ierror Then Exit Sub

End Sub

Private Sub Command2_Click()
If Not DebugMode Then On Error Resume Next
Unload FormSETCMP
DoEvents
End Sub

Private Sub Command3_Click()
If Not DebugMode Then On Error Resume Next
Call GetCmpSetCmpClear
If ierror Then Exit Sub
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormSETCMP)
HelpContextID = IOGetHelpContextID("FormSETCMP")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub TextCharge_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextComposition_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

