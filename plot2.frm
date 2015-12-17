VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FormPLOT 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plot!"
   ClientHeight    =   4350
   ClientLeft      =   1530
   ClientTop       =   1470
   ClientWidth     =   5730
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
   Icon            =   "PLOT2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4350
   ScaleWidth      =   5730
   Begin ComctlLib.StatusBar StatusBarPlot 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4095
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   6456
            TextSave        =   ""
            Key             =   "status"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Plot status"
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
            Object.ToolTipText     =   "Click here to cancel the plot"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Bevel           =   2
            Enabled         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Next"
            TextSave        =   "Next"
            Key             =   "next"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Click here to plot the next sample"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Stub Form"
      Height          =   735
      Left            =   1560
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   2415
   End
End
Attribute VB_Name = "FormPLOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2015 by John J. Donovan
Option Explicit

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormPLOT)
HelpContextID = IOGetHelpContextID("FormPLOT")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub StatusBarPlot_PanelClick(ByVal Panel As ComctlLib.Panel)
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
