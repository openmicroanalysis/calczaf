VERSION 5.00
Begin VB.Form FormHISTO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Binary Histogram Options"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Histogram Options"
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
      Height          =   2655
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3975
      Begin VB.TextBox TextHistogramNumberofBuckets 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox TextHistogramMaximum 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox TextHistogramMinimum 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Histogram Number of ""Buckets"""
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
         TabIndex        =   8
         Top             =   1920
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "Histogram Maximum Value"
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
         TabIndex        =   7
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Histogram Minimum Value"
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
         TabIndex        =   6
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command2 
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
      Left            =   4200
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "FormHISTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2016 by John J. Donovan
Option Explicit

Private Sub Command1_Click()
' User clicked OK in form HISTO
If Not DebugMode Then On Error Resume Next
' Save the options
Call CalcZAFHistogramSave
If ierror Then Exit Sub
Unload FormHISTO
DoEvents
' Plot histogram
Call CalcZAFPlotHistogram(Int(0))
If ierror Then Exit Sub
End Sub

Private Sub Command2_Click()
' User clicked Cancel in form HISTO
If Not DebugMode Then On Error Resume Next
Unload FormHISTO
DoEvents
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormHISTO)
HelpContextID = IOGetHelpContextID("FormHISTO")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub TextHistogramMaximum_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextHistogramMinimum_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextHistogramNumberofBuckets_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

