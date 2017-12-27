VERSION 5.00
Begin VB.Form FormUnzip 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Unzip Files From ZIP"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton CommandExtract 
      Caption         =   "Extract"
      Height          =   495
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton CommandClose 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   495
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox TextUnzipFolder 
      Height          =   405
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   4335
   End
   Begin VB.TextBox TextUnzipFile 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "Folder To UnZIP To"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "File To UnZIP"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "FormUnzip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2018 by John J. Donovan
Option Explicit

Dim tpassword As String

Private WithEvents m_cUnzip As cUnzip
Attribute m_cUnzip.VB_VarHelpID = -1

Private Sub CommandClose_Click()
Unload Me
End Sub

Public Sub CommandExtract_Click()
' Call this to extract files

ierror = False
On Error GoTo CommandExtractError

Dim tfilename As String
Dim tfolder As String

tfilename$ = FormUnzip.TextUnzipFile.Text
tfolder$ = FormUnzip.TextUnzipFolder.Text

If DebugMode Then FormUnzip.Show vbModeless

' Set ZIP file
m_cUnzip.ZipFile = tfilename$

' Set base folder to unzip to
m_cUnzip.UnzipFolder = tfolder$

m_cUnzip.OverwriteExisting = True
m_cUnzip.UseFolderNames = True

' Unzip the file
m_cUnzip.Unzip

Exit Sub

' Errors
CommandExtractError:
MsgBox Error$, vbOKOnly + vbCritical, "CommandExtract"
ierror = True
Exit Sub

End Sub

Private Sub Form_Load()
' Set up unzipping object
Set m_cUnzip = New cUnzip

Call MiscLoadIcon(FormUnzip)
Call MiscCenterForm(FormUnzip)
HelpContextID = IOGetHelpContextID("FormUnzip")
End Sub

Private Sub Form_Unload(Cancel As Integer)
' Important: to ensure this class terminates we must set to nothing here:
Set m_cUnzip = Nothing
tpassword$ = vbNullString
End Sub

Private Sub m_cUnzip_Progress(ByVal lCount As Long, ByVal sMsg As String)
Call IOWriteLog(vbCrLf & sMsg$)
DoEvents
End Sub

Private Sub m_cUnzip_PasswordRequest(sPassword As String, bCancel As Boolean)
If tpassword$ <> vbNullString Then
sPassword$ = tpassword$
Exit Sub
End If
Call PasswordLoad(sPassword$, "Enter ZIP Password", "Enter the ZIP file password for the specified ZIP file.")
If ierror Then Exit Sub
tpassword$ = sPassword$
End Sub

