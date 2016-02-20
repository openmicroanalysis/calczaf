VERSION 5.00
Begin VB.Form FormREGISTER 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Software Registration"
   ClientHeight    =   6360
   ClientLeft      =   1605
   ClientTop       =   2340
   ClientWidth     =   5745
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
   ScaleHeight     =   6360
   ScaleWidth      =   5745
   Begin VB.CommandButton CommandCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton CommandHelp 
      Caption         =   "Help"
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
      Left            =   4320
      TabIndex        =   7
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton CommandRegister 
      BackColor       =   &H0000FFFF&
      Caption         =   "Register"
      Default         =   -1  'True
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Registration Information"
      ForeColor       =   &H000000FF&
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.TextBox TextInstitution 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   3855
      End
      Begin VB.TextBox TextName 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   3855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   $"REGISTER.frx":0000
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
         Height          =   855
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "Company, School or Institution"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   3855
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "Customer or User Name"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   3855
      End
   End
   Begin VB.Label LabelDisclaimer 
      Height          =   3375
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   5535
   End
End
Attribute VB_Name = "FormREGISTER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2016 by John J. Donovan
Option Explicit

Private Sub CommandCancel_Click()
If Not DebugMode Then On Error Resume Next
End
End Sub

Private Sub CommandHelp_Click()
If Not DebugMode Then On Error Resume Next
Call MiscFormLoadHelp(FormREGISTER.HelpContextID)
If ierror Then Exit Sub
End Sub

Private Sub CommandRegister_Click()
If Not DebugMode Then On Error Resume Next
Call RegisterSave
If ierror Then Exit Sub
Unload FormREGISTER
End Sub

Private Sub Form_Activate()
' Load FormREGISTER text
If Not DebugMode Then On Error Resume Next
Dim tmsg As String

tmsg$ = "IN NO EVENT SHALL PROBE SOFTWARE BE LIABLE TO ANY PARTY "
tmsg$ = tmsg$ & "FOR DIRECT, INDIRECT, SPECIAL, INICIDENTAL, OR CONSEQUENTIAL DAMAGES, "
tmsg$ = tmsg$ & "INCLUDING LOST PROFITS, ARISING OUT OF THE USE OF THIS SOFTWARE AND ITS "
tmsg$ = tmsg$ & "DOCUMENTATION, EVEN IF PROBE SOFTWARE HAS BEEN ADVISED OF "
tmsg$ = tmsg$ & "THE POSSIBILITY OF SUCH DAMAGE." & vbCrLf & vbCrLf

tmsg$ = tmsg$ & "PROBE SOFTWARE SPECIFICALLY DISCLAIMS ANY WARRANTIES, "
tmsg$ = tmsg$ & "INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF "
tmsg$ = tmsg$ & "MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE. THE SOFTWARE "
tmsg$ = tmsg$ & "PROVIDED  HEREUNDER IS ON AN AS IS BASIS, AND PROBE SOFTWARE "
tmsg$ = tmsg$ & "HAVE NO OBLIGATIONS TO PROVIDE MAINTENANCE, SUPPORT, UPDATES, "
tmsg$ = tmsg$ & "ENHANCEMENTS, OR MODIFICATIONS."

FormREGISTER.LabelDisclaimer.Caption = tmsg$
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call MiscCenterForm(FormREGISTER)
Call MiscLoadIcon(FormREGISTER)
HelpContextID = IOGetHelpContextID("FormREGISTER")
End Sub

Private Sub TextInstitution_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextName_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

