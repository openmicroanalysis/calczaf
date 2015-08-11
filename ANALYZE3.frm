VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FormANALYZE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Analyze!"
   ClientHeight    =   5475
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   6495
   Icon            =   "ANALYZE3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.OptionButton OptionStandard 
      Caption         =   "Standards"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1335
   End
   Begin VB.OptionButton OptionUnknown 
      Caption         =   "Unknowns"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1335
   End
   Begin VB.OptionButton OptionAllSamples 
      Caption         =   "All Samples"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1335
   End
   Begin VB.OptionButton OptionWavescan 
      Caption         =   "Wavescans"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1335
   End
   Begin VB.ListBox ListAnalyze 
      Height          =   2010
      Left            =   3360
      MultiSelect     =   2  'Extended
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   2895
   End
   Begin VB.CheckBox CheckDoNotOutputToLog 
      Caption         =   "Do Not Output To Log"
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton CommandNext 
      Height          =   255
      Left            =   3600
      Picture         =   "ANALYZE3.frx":6E7FA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin ComctlLib.StatusBar StatusBarAnal 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5220
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   7329
            TextSave        =   ""
            Key             =   "status"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Cancel"
            TextSave        =   "Cancel"
            Key             =   "cancel"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Next"
            TextSave        =   "Next"
            Key             =   "next"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stub Form"
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
      Left            =   1800
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   2415
   End
End
Attribute VB_Name = "FormANALYZE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2015 by John J. Donovan
Option Explicit

Private Sub StatusBarAnal_PanelClick(ByVal Panel As ComctlLib.Panel)
If Not DebugMode Then On Error Resume Next
Select Case Panel.key
Case "status"
    Exit Sub
Case "cancel"
    Call AnalyzeCancel(FormANALYZE)
    If ierror Then Exit Sub
Case "next"
    Call AnalyzeNext(Int(2))
    If ierror Then Exit Sub
Case Else
    Exit Sub
End Select
End Sub
