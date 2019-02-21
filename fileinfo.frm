VERSION 5.00
Begin VB.Form FormFILEINFO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Information"
   ClientHeight    =   4665
   ClientLeft      =   2340
   ClientTop       =   1620
   ClientWidth     =   8745
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
   ScaleHeight     =   4665
   ScaleWidth      =   8745
   Begin VB.CommandButton CommandAddCR 
      Caption         =   "Add <cr>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   26
      ToolTipText     =   "Add a carriage return to the description text (place cursor and hit Add <cr>)"
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox TextCustom3 
      Height          =   285
      Left            =   6360
      TabIndex        =   4
      ToolTipText     =   "User defined custom field #3"
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox TextCustom2 
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      ToolTipText     =   "User defined custom field #2"
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox TextCustom1 
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      ToolTipText     =   "User defined custom field #1"
      Top             =   1680
      Width           =   5175
   End
   Begin VB.TextBox TextUser 
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      ToolTipText     =   "The user name"
      Top             =   960
      Width           =   5175
   End
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7560
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton CommandOK 
      BackColor       =   &H00C0FFC0&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox TextDescription 
      Height          =   1455
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      ToolTipText     =   "Database file description"
      Top             =   2400
      Width           =   6375
   End
   Begin VB.TextBox TextTitle 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      ToolTipText     =   "The title of the probe database run"
      Top             =   1320
      Width           =   5175
   End
   Begin VB.Label LabelDateModified 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6360
      TabIndex        =   25
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label LabelLastUpdated 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2280
      TabIndex        =   24
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label LabelDateCreated 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2280
      TabIndex        =   23
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label LabelType 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5760
      TabIndex        =   22
      ToolTipText     =   "The type of the database (STANDARD, PROBE, SETUP, USER, POSITION and XRAY)"
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label LabelVersion 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2280
      TabIndex        =   21
      ToolTipText     =   "The version number of the database when it was created"
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label LabelFileName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2280
      TabIndex        =   20
      ToolTipText     =   "The database path and filename "
      Top             =   120
      Width           =   6375
   End
   Begin VB.Label LabelCustom3 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4200
      TabIndex        =   19
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label LabelCustom2 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label LabelCustom1 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      Caption         =   "File Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      Caption         =   "Date Modified"
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
      Height          =   255
      Left            =   4920
      TabIndex        =   15
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      Caption         =   "User"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      Caption         =   "Last Updated"
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
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      Caption         =   "Date Created"
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
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      Caption         =   "Description"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      Caption         =   "Title"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "Type"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5040
      TabIndex        =   7
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "Version"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "FormFILEINFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2019 by John J. Donovan
Option Explicit

Private Sub CommandAddCR_Click()
If Not DebugMode Then On Error Resume Next
Call MiscAddCRToText(FormFILEINFO.TextDescription)
If ierror Then Exit Sub
End Sub

Private Sub CommandCancel_Click()
If Not DebugMode Then On Error Resume Next
Unload FormFILEINFO
icancelload = True
End Sub

Private Sub CommandOK_Click()
If Not DebugMode Then On Error Resume Next
Call FileInfoSave
If ierror Then Exit Sub
Unload FormFILEINFO
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
icancelload = False
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormFILEINFO)
HelpContextID = IOGetHelpContextID("FormFILEINFO")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub TextCustom1_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextCustom2_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextCustom3_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextDescription_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextTitle_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextUser_Change()
If Not DebugMode Then On Error Resume Next

Dim tmsg As String

If RealTimeMode Then
If UCase$(app.EXEName) = UCase$("PROBEWIN") Then
tmsg$ = FormFILEINFO.TextUser.Text
Call UserGetLastRecord(tmsg$, FormFILEINFO)
If ierror Then Exit Sub
End If
End If

End Sub

Private Sub TextUser_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

