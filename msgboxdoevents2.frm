VERSION 5.00
Begin VB.Form FormMSGBOXDOEVENTS2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MsgBoxDoEvents"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4680
      Top             =   1320
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
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Version for real time processing of driver without interface calls"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   975
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "MsgBoxDoevents2.frx":0000
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "FormMSGBOXDOEVENTS2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2017 by John J. Donovan
Option Explicit

Private Sub CommandOK_Click()
If Not DebugMode Then On Error Resume Next
FormMSGBOXDOEVENTS2.Timer1.Enabled = False
DoEvents
Unload FormMSGBOXDOEVENTS2
End Sub

Private Sub Form_Activate()
If Not DebugMode Then On Error Resume Next
FormMSGBOXDOEVENTS2.Timer1.Enabled = True
DoEvents
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call MiscAlwaysOnTop(True, FormMSGBOXDOEVENTS2)
Call MiscCenterForm(FormMSGBOXDOEVENTS2)
Call MiscLoadIcon(FormMSGBOXDOEVENTS2)
HelpContextID = IOGetHelpContextID("FormMSGBOXDOEVENTS2")
End Sub

Private Sub Timer1_Timer()
If Not DebugMode Then On Error Resume Next
DoEvents    ' allow other processes to run
Sleep 0     ' allow other applications to run
End Sub
