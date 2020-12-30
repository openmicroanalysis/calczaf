VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Object = "{0EF1F64B-CADD-4E29-8703-556548DF74E3}#1.0#0"; "cshtpctl.ocx"
Object = "{3A1209F5-3069-4E0A-A192-1427CFD1D5A9}#1.0#0"; "csftpctl.ocx"
Begin VB.Form FormUPDATE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Probe for EPMA"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin FtpClientCtl.FtpClient FtpClient1 
      Left            =   2640
      Top             =   720
      _cx             =   741
      _cy             =   741
   End
   Begin HttpClientCtl.HttpClient HttpClient1 
      Left            =   3120
      Top             =   720
      _cx             =   741
      _cy             =   741
   End
   Begin VB.CheckBox CheckUpdatePenepmaOnly 
      Caption         =   "Update Penepma Monte Carlo Files Only"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   2040
      Width           =   4095
   End
   Begin VB.OptionButton OptionDownloadType 
      Caption         =   "Download Update Using Alternative FTP"
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
      Index           =   2
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Download from ""probesoftware"""
      Top             =   1680
      Width           =   3855
   End
   Begin VB.CommandButton CommandUpdate 
      BackColor       =   &H0080FFFF&
      Caption         =   "Download Update!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Download the latest update file. The program will ask if you want to automatically extract the downloaded update."
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton CommandDeleteUpdate 
      Caption         =   "Delete Update"
      Height          =   495
      Left            =   4080
      TabIndex        =   6
      ToolTipText     =   "Delete the previous update file download (to force an update)"
      Top             =   720
      Width           =   1095
   End
   Begin VB.OptionButton OptionDownloadType 
      Caption         =   "Download Update Using HTTP Protocol"
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
      Index           =   1
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Download from ""epmalab"""
      Top             =   1200
      Value           =   -1  'True
      Width           =   3855
   End
   Begin VB.OptionButton OptionDownloadType 
      Caption         =   "Download Update Using FTP Protocol"
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
      Index           =   0
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Download from ""whitewater"""
      Top             =   1440
      Width           =   3855
   End
   Begin VB.CommandButton CommandClose 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Close"
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label LabelInstructions 
      Caption         =   $"Update.frx":0000
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "FormUPDATE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2021 by John J. Donovan
Option Explicit

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
Unload FormUPDATE
End Sub

Private Sub CommandDeleteUpdate_Click()
If Not DebugMode Then On Error Resume Next
Call IOUpdateDeleteUpdate
End Sub

Private Sub CommandUpdate_Click()
If Not DebugMode Then On Error Resume Next
If FormUPDATE.OptionDownloadType(0).value = True Then
FormUPDATE.OptionDownloadType(0).Enabled = False
FormUPDATE.OptionDownloadType(1).Enabled = False
FormUPDATE.OptionDownloadType(2).Enabled = False
FormUPDATE.CommandUpdate.Enabled = False
FormUPDATE.CheckUpdatePenepmaOnly.Enabled = False
Call IOUpdateGetUpdate(Int(1))
FormUPDATE.CommandUpdate.Enabled = True
FormUPDATE.CheckUpdatePenepmaOnly.Enabled = True
FormUPDATE.OptionDownloadType(0).Enabled = True
FormUPDATE.OptionDownloadType(1).Enabled = True
FormUPDATE.OptionDownloadType(2).Enabled = True
If ierror Then Exit Sub

ElseIf FormUPDATE.OptionDownloadType(1).value = True Then
FormUPDATE.OptionDownloadType(0).Enabled = False
FormUPDATE.OptionDownloadType(1).Enabled = False
FormUPDATE.OptionDownloadType(2).Enabled = False
FormUPDATE.CommandUpdate.Enabled = False
FormUPDATE.CheckUpdatePenepmaOnly.Enabled = False
Call IOUpdateGetUpdate(Int(2))
FormUPDATE.CommandUpdate.Enabled = True
FormUPDATE.CheckUpdatePenepmaOnly.Enabled = True
FormUPDATE.OptionDownloadType(0).Enabled = True
FormUPDATE.OptionDownloadType(1).Enabled = True
FormUPDATE.OptionDownloadType(2).Enabled = True
If ierror Then Exit Sub

Else
FormUPDATE.OptionDownloadType(0).Enabled = False
FormUPDATE.OptionDownloadType(1).Enabled = False
FormUPDATE.OptionDownloadType(2).Enabled = False
FormUPDATE.CommandUpdate.Enabled = False
FormUPDATE.CheckUpdatePenepmaOnly.Enabled = False
Call IOUpdateGetUpdate(Int(3))
FormUPDATE.CommandUpdate.Enabled = True
FormUPDATE.CheckUpdatePenepmaOnly.Enabled = True
FormUPDATE.OptionDownloadType(0).Enabled = True
FormUPDATE.OptionDownloadType(1).Enabled = True
FormUPDATE.OptionDownloadType(2).Enabled = True
If ierror Then Exit Sub
End If
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormUPDATE)
HelpContextID = IOGetHelpContextID("FormUPDATE")

If UCase$(app.EXEName) = UCase$("CalcZAF") Then
FormUPDATE.Caption = "Update CalcZAF"
FormUPDATE.LabelInstructions.Caption = "Make sure you are connected to the Internet, select the download type (FTP or HTTP)  and click the Download Update button to download the CalcZAF update"
End If

If UCase$(app.EXEName) = UCase$("Probewin") Then
FormUPDATE.Caption = "Update Probe for EPMA"
FormUPDATE.LabelInstructions.Caption = "Make sure you are connected to the Internet, select the download type (FTP or HTTP)  and click the Download Update button to download the Probe for EPMA update"
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
Call IOUpdateClose(Cancel%)
If ierror Then Exit Sub
End Sub

Private Sub FtpClient1_OnProgress(ByVal filename As Variant, ByVal FileSize As Variant, ByVal BytesCopied As Variant, ByVal Percent As Variant)
If Not DebugMode Then On Error Resume Next
DoEvents
FormUPDATE.ProgressBar1.value = Percent
If UCase$(app.EXEName) = UCase$("CalcZAF") Then
If Percent > 0# Then FormUPDATE.Caption = "Update CalcZAF [" & Format$(Percent) & "% downloaded...]"
End If

If UCase$(app.EXEName) = UCase$("Probewin") Then
If Percent > 0# Then FormUPDATE.Caption = "Update Probe for EPMA [" & Format$(Percent) & "% downloaded...]"
End If
DoEvents
End Sub

Private Sub HttpClient1_OnProgress(ByVal BytesTotal As Variant, ByVal BytesCopied As Variant, ByVal Percent As Variant)
If Not DebugMode Then On Error Resume Next
DoEvents
FormUPDATE.ProgressBar1.value = Percent
If UCase$(app.EXEName) = UCase$("CalcZAF") Then
If Percent > 0# Then FormUPDATE.Caption = "Update CalcZAF [" & Format$(Percent) & "% downloaded...]"
End If

If UCase$(app.EXEName) = UCase$("Probewin") Then
If Percent > 0# Then FormUPDATE.Caption = "Update Probe for EPMA [" & Format$(Percent) & "% downloaded...]"
End If
DoEvents
End Sub

