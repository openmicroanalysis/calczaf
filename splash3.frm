VERSION 5.00
Begin VB.Form FormSPLASH 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   4185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8610
   ControlBox      =   0   'False
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "SPLASH3.frx":0000
   ScaleHeight     =   4185
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   1920
   End
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   120
      Top             =   1320
   End
   Begin VB.Image Image5 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      Picture         =   "SPLASH3.frx":75726
      Stretch         =   -1  'True
      Top             =   2510
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image3 
      Height          =   1230
      Left            =   1170
      Picture         =   "SPLASH3.frx":773F2
      Top             =   1680
      Width           =   5940
   End
   Begin VB.Image Image4 
      Height          =   100
      Left            =   2100
      Picture         =   "SPLASH3.frx":77EFA
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   5505
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "EPMA Calculation and Modeling Utility"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1080
      Width           =   6015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CalcZAF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   6000
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   960
      Left            =   120
      Picture         =   "SPLASH3.frx":78B24
      Top             =   240
      Width           =   960
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Distributed by Probe Software, Inc."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3600
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) 1995-2021 by John J. Donovan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   3240
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   1185
      Left            =   6120
      Picture         =   "SPLASH3.frx":79225
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   2370
   End
End
Attribute VB_Name = "FormSPLASH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2021 by John J. Donovan
Option Explicit

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call MiscCenterForm(FormSPLASH)
FormSPLASH.Image5.Visible = True
FormSPLASH.Timer2.Enabled = True
End Sub

Private Sub Form_Resize()
If Not DebugMode Then On Error Resume Next
FormSPLASH.Image1.Move 0, 0, FormSPLASH.ScaleWidth, FormSPLASH.ScaleHeight
End Sub

Private Sub Timer1_Timer()
If Not DebugMode Then On Error Resume Next
Unload FormSPLASH
End Sub

Private Sub Timer2_Timer()
FormSPLASH.Image5.Left = FormSPLASH.Image5.Left + 80
If FormSPLASH.Image5.Left > (FormSPLASH.Image4.Left + FormSPLASH.Image4.Width) - 300 Then
FormSPLASH.Image5.Visible = False
FormSPLASH.Timer2.Enabled = False
End If
End Sub
