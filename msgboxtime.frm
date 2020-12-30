VERSION 5.00
Begin VB.Form FormMSGBOXTIME 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MsgBoxTime"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label1 
      Height          =   2175
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   7095
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "MsgBoxTime.frx":0000
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "FormMSGBOXTIME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2021 by John J. Donovan
Option Explicit
' This message box has only an OK button
Private Sub CommandOK_Click()
If Not DebugMode Then On Error Resume Next
Unload FormMSGBOXTIME
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call MiscCenterForm(FormMSGBOXTIME)
Call MiscLoadIcon(FormMSGBOXTIME)
HelpContextID = IOGetHelpContextID("FormMSGBOXTIME")
End Sub

Private Sub Timer1_Timer()
If Not DebugMode Then On Error Resume Next
Unload FormMSGBOXTIME
End Sub
