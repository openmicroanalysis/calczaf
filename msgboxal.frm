VERSION 5.00
Begin VB.Form FormMSGBOXALL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MsgBoxAll"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   120
      Top             =   1080
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
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton CommandYesToAll 
      Caption         =   "Yes To All"
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
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton CommandNo 
      Caption         =   "No"
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
      Left            =   1200
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton CommandYes 
      Caption         =   "Yes"
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
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   1455
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "MsgBoxAl.frx":0000
      Top             =   480
      Width           =   480
   End
End
Attribute VB_Name = "FormMSGBOXALL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2022 by John J. Donovan
Option Explicit

Private Sub CommandCancel_Click()
If Not DebugMode Then On Error Resume Next
MsgBoxAllReturnValue% = 2
Unload FormMSGBOXALL
End Sub

Private Sub CommandNo_Click()
If Not DebugMode Then On Error Resume Next
MsgBoxAllReturnValue% = 7
Unload FormMSGBOXALL
End Sub

Private Sub CommandYes_Click()
If Not DebugMode Then On Error Resume Next
MsgBoxAllReturnValue% = 6
Unload FormMSGBOXALL
End Sub

Private Sub CommandYesToAll_Click()
If Not DebugMode Then On Error Resume Next
MsgBoxAllReturnValue% = 8
Unload FormMSGBOXALL
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call MiscCenterForm(FormMSGBOXALL)
Call MiscLoadIcon(FormMSGBOXALL)
HelpContextID = IOGetHelpContextID("FormMSGBOXALL")
End Sub

Private Sub Timer1_Timer()
If Not DebugMode Then On Error Resume Next
MsgBoxAllReturnValue% = 8       ' assume yes to all if timer elapses
Unload FormMSGBOXALL
End Sub
