VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Object = "{58B27513-14ED-4580-8D51-6739523755CC}#1.0#0"; "cshtpx11.ocx"
Begin VB.Form FormUPDATE2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Probe for EPMA"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HttpClientCtl.HttpClient HttpClient1 
      Left            =   3600
      Top             =   120
      _cx             =   741
      _cy             =   741
   End
   Begin VB.CheckBox CheckUpdatePenepmaOnly 
      Caption         =   "Update Penepma Monte Carlo Files Only"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1800
      Width           =   3255
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
      Caption         =   "Download update from Probe Software"
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
      ToolTipText     =   "Download from probesoftware.com using https"
      Top             =   1080
      Value           =   -1  'True
      Width           =   3855
   End
   Begin VB.OptionButton OptionDownloadType 
      Caption         =   "Download update from UofO EPMA Lab"
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
      ToolTipText     =   "Download from epmalab.uoregon.edu using https"
      Top             =   1320
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
      Top             =   2160
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label LabelInstructions 
      Caption         =   $"Update2.frx":0000
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "FormUPDATE2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2026 by John J. Donovan
Option Explicit

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
Unload FormUPDATE2
End Sub

Private Sub CommandDeleteUpdate_Click()
If Not DebugMode Then On Error Resume Next
Call IOUpdate2DeleteUpdate
End Sub

Private Sub CommandUpdate_Click()
If Not DebugMode Then On Error Resume Next
If FormUPDATE2.OptionDownloadType(0).value = True Then      ' UofO EPMA Lab
FormUPDATE2.OptionDownloadType(0).Enabled = False
FormUPDATE2.OptionDownloadType(1).Enabled = False
FormUPDATE2.CommandUpdate.Enabled = False
FormUPDATE2.CheckUpdatePenepmaOnly.Enabled = False
Call IOUpdate2GetUpdate(Int(1))
FormUPDATE2.CommandUpdate.Enabled = True
FormUPDATE2.CheckUpdatePenepmaOnly.Enabled = True
FormUPDATE2.OptionDownloadType(0).Enabled = True
FormUPDATE2.OptionDownloadType(1).Enabled = True
If ierror Then Exit Sub

ElseIf FormUPDATE2.OptionDownloadType(1).value = True Then  ' Probe Software (default)
FormUPDATE2.OptionDownloadType(0).Enabled = False
FormUPDATE2.OptionDownloadType(1).Enabled = False
FormUPDATE2.CommandUpdate.Enabled = False
FormUPDATE2.CheckUpdatePenepmaOnly.Enabled = False
Call IOUpdate2GetUpdate(Int(2))
FormUPDATE2.CommandUpdate.Enabled = True
FormUPDATE2.CheckUpdatePenepmaOnly.Enabled = True
FormUPDATE2.OptionDownloadType(0).Enabled = True
FormUPDATE2.OptionDownloadType(1).Enabled = True
If ierror Then Exit Sub
End If
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormUPDATE2)
HelpContextID = IOGetHelpContextID("FormUpdate2")

If UCase$(app.EXEName) = UCase$("CalcZAF") Then
FormUPDATE2.Caption = "Update CalcZAF"
FormUPDATE2.LabelInstructions.Caption = "Make sure you are connected to the Internet, select the download site and click the Download Update button to download the CalcZAF (or PENEPMA12) update"
End If

If UCase$(app.EXEName) = UCase$("Probewin") Then
FormUPDATE2.Caption = "Update Probe for EPMA"
FormUPDATE2.LabelInstructions.Caption = "Make sure you are connected to the Internet, select the download site and click the Download Update button to download the Probe for EPMA (or PENEPMA) update"
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
Call IOUpdate2Close(Cancel%)
If ierror Then Exit Sub
End Sub

Private Sub FtpClient1_OnProgress(ByVal filename As Variant, ByVal FileSize As Variant, ByVal BytesCopied As Variant, ByVal Percent As Variant)
If Not DebugMode Then On Error Resume Next
DoEvents
FormUPDATE2.ProgressBar1.value = Percent
If UCase$(app.EXEName) = UCase$("CalcZAF") Then
If Percent > 0# Then FormUPDATE2.Caption = "Update CalcZAF [" & Format$(Percent) & "% downloaded...]"
End If

If UCase$(app.EXEName) = UCase$("Probewin") Then
If Percent > 0# Then FormUPDATE2.Caption = "Update Probe for EPMA [" & Format$(Percent) & "% downloaded...]"
End If
DoEvents
End Sub

Private Sub HttpClient1_OnProgress(ByVal BytesTotal As Variant, ByVal BytesCopied As Variant, ByVal Percent As Variant)
If Not DebugMode Then On Error Resume Next
DoEvents
FormUPDATE2.ProgressBar1.value = Percent
If UCase$(app.EXEName) = UCase$("CalcZAF") Then
If Percent > 0# Then FormUPDATE2.Caption = "Update CalcZAF [" & Format$(Percent) & "% downloaded...]"
End If

If UCase$(app.EXEName) = UCase$("Probewin") Then
If Percent > 0# Then FormUPDATE2.Caption = "Update Probe for EPMA [" & Format$(Percent) & "% downloaded...]"
End If
DoEvents
End Sub

