VERSION 5.00
Object = "{6E5043E8-C452-4A6A-B011-9B5687112610}#1.0#0"; "Pesgo32f.ocx"
Begin VB.Form FormPICTURESNAP4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PictureSnap Digitize Polygon"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TextSampleName 
      Height          =   285
      Left            =   480
      TabIndex        =   8
      ToolTipText     =   "Enter the sample name"
      Top             =   5760
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Export Digitized Coordinates"
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
      Left            =   480
      TabIndex        =   7
      Top             =   6120
      Width           =   3375
   End
   Begin Pesgo32fLib.Pesgo Pesgo1 
      Height          =   4095
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   4095
      _Version        =   65536
      _ExtentX        =   7223
      _ExtentY        =   7223
      _StockProps     =   96
      _AllProps       =   "PICTURESNAP4.frx":0000
   End
   Begin VB.CommandButton CommandDigitizeClose 
      Caption         =   "Close Digitize"
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
      Left            =   2400
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton CommandDigitizeStart 
      Caption         =   "Start Digitize"
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
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton CommandClose 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Close"
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
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Beam or Stage Scan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.OptionButton OptionBeamOrStage 
         Caption         =   "Stage Scan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   3
         Top             =   420
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptionBeamOrStage 
         Caption         =   "Beam Scan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   420
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FormPICTURESNAP4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2017 by John J. Donovan
Option Explicit

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
Unload FormPICTURESNAP4
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormPICTURESNAP4)
HelpContextID = IOGetHelpContextID("FormPICTURESNAP4")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub TextSampleName_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub
