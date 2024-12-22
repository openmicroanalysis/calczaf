VERSION 5.00
Begin VB.Form FormMSGBOXTIM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MsgBoxTim"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   7950
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
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   1920
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   120
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label1 
      Height          =   1695
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   7095
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "MsgBoxTi.frx":0000
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "FormMSGBOXTIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2025 by John J. Donovan
Option Explicit
' This message box has a OK and Cancel button
Private Sub CommandCancel_Click()
If Not DebugMode Then On Error Resume Next
Unload FormMSGBOXTIM
ierror = True
End Sub

Private Sub CommandOK_Click()
If Not DebugMode Then On Error Resume Next
Unload FormMSGBOXTIM
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call MiscCenterForm(FormMSGBOXTIM)
Call MiscLoadIcon(FormMSGBOXTIM)
HelpContextID = IOGetHelpContextID("FormMSGBOXTIM")
End Sub

Private Sub Timer1_Timer()
If Not DebugMode Then On Error Resume Next
Unload FormMSGBOXTIM
End Sub
