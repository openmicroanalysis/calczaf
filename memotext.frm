VERSION 5.00
Begin VB.Form FormMemoText 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Memo Text for Standard"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CommandEditSelectAll 
      Caption         =   "Select All"
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
      Left            =   9960
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton CommandEditPaste 
      Caption         =   "Paste"
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
      Left            =   9960
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton CommandEditCut 
      Caption         =   "Cut"
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
      Left            =   9960
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton CommandEditCopy 
      Caption         =   "Copy"
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
      Left            =   9960
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton CommandEditClearAll 
      Caption         =   "Clear All"
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
      Left            =   9960
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton CommandOK 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "OK"
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
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton CommandCancel 
      Appearance      =   0  'Flat
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
      Left            =   9720
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox TextMemoText 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   9495
   End
End
Attribute VB_Name = "FormMemoText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2020 by John J. Donovan
Option Explicit

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormMemoText)
HelpContextID = IOGetHelpContextID("FormMEMOTEXT")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub CommandCancel_Click()
If Not DebugMode Then On Error Resume Next
Unload FormMemoText
ierror = True
End Sub

Private Sub CommandOK_Click()
If Not DebugMode Then On Error Resume Next
Call GetCmpMemoTextSave
If ierror Then Exit Sub
End Sub

Private Sub CommandEditClearAll_Click()
If Not DebugMode Then On Error Resume Next
FormMemoText.TextMemoText.Text = vbNullString
End Sub

Private Sub CommandEditCopy_Click()
If Not DebugMode Then On Error Resume Next
Clipboard.Clear
Clipboard.SetText FormMemoText.TextMemoText.SelText
End Sub

Private Sub CommandEditCut_Click()
If Not DebugMode Then On Error Resume Next
Clipboard.Clear
Clipboard.SetText FormMemoText.TextMemoText.SelText
FormMemoText.TextMemoText.SelText = vbNullString
End Sub

Private Sub CommandEditPaste_Click()
If Not DebugMode Then On Error Resume Next
FormMemoText.TextMemoText.SelText = Clipboard.GetText()
End Sub

Private Sub CommandEditSelectAll_Click()
If Not DebugMode Then On Error Resume Next
FormMemoText.TextMemoText.SetFocus
FormMemoText.TextMemoText.SelStart = 0
FormMemoText.TextMemoText.SelLength = Len(FormMemoText.TextMemoText.Text)
End Sub


