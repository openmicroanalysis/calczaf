VERSION 5.00
Begin VB.Form FormSECONDARYKratios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Specify K-Ratios.DAT file for Secondary Fluorescence From Boundary"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
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
      Height          =   615
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   1095
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
      Height          =   615
      Left            =   8280
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      Caption         =   "Specify FANAL K-Ratios From Previously Calculated PAR File Couple"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   9255
      Begin VB.CheckBox CheckUseSecondaryFluorescenceCorrection 
         Caption         =   "Perform Secondary Fluorescence Boundary Correction"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2760
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Flag to specify secondary boundary fluoresence corerction for this element"
         Top             =   480
         Width           =   5055
      End
      Begin VB.CommandButton CommandHelp 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Help"
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
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Click this button to get detailed help from our on-line user forum"
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CommandBrowseForCouple 
         BackColor       =   &H0080FFFF&
         Caption         =   "Browse For SF Couple"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   $"SecondaryKratios.frx":0000
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label LabelKratiosDATFile 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   8775
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   $"SecondaryKratios.frx":00DC
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "FormSECONDARYKratios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2026 by John J. Donovan
Option Explicit

Private Sub CommandBrowseForCouple_Click()
If Not DebugMode Then On Error Resume Next
Call SecondaryBrowseFile(Int(0), FormSECONDARYKratios)
If ierror Then
msg$ = "There was an error reading the K-ratio2.dat file." & vbCrLf & vbCrLf
msg$ = msg$ & "Make sure the DAT file is in a folder which is named so it contains the kilovolts, beam incident, boundary and standard materials, and element atomic number and x-ray line (e.g., measuring Si Ka in SiO2 adjacent to TiO2 using tiO2 as a standard at 15 keV would be 15_SiO2_TiO2_TiO2_22_1)."
MsgBox msg$, vbOKOnly + vbExclamation, "FormSECONDARYKratios"
Exit Sub
End If
End Sub

Private Sub CommandCancel_Click()
If Not DebugMode Then On Error Resume Next
Unload FormSECONDARYKratios
End Sub

Private Sub CommandHelp_Click()
If Not DebugMode Then On Error Resume Next
Call IOBrowseHTTP(ProbeSoftwareInternetBrowseMethod%, "https://smf.probesoftware.com/index.php?topic=58.msg214#msg214")
If ierror Then Exit Sub
End Sub

Private Sub CommandOK_Click()
If Not DebugMode Then On Error Resume Next
Call SecondaryKratiosSave
If ierror Then Exit Sub
Unload FormSECONDARYKratios
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormSECONDARYKratios)
HelpContextID = IOGetHelpContextID("FormSECONDARYKratios")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

