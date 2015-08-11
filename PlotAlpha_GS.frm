VERSION 5.00
Object = "{827E9F53-96A4-11CF-823E-000021570103}#1.0#0"; "graphs32.ocx"
Begin VB.Form FormPlotAlpha_GS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calculate and Plot Binary Alpha Factors"
   ClientHeight    =   8355
   ClientLeft      =   900
   ClientTop       =   900
   ClientWidth     =   10800
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
   ScaleHeight     =   8355
   ScaleWidth      =   10800
   Begin VB.OptionButton OptionBenceAlbee 
      Caption         =   "Option1"
      Height          =   615
      Index           =   3
      Left            =   8760
      TabIndex        =   21
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton CommandSaveData 
      Caption         =   "Export Data"
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
      Left            =   8760
      TabIndex        =   18
      Top             =   7440
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Zoom Full"
      Height          =   375
      Left            =   8760
      TabIndex        =   4
      Top             =   7920
      Width           =   1815
   End
   Begin VB.CommandButton CommandPrint 
      Caption         =   "Print"
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
      Left            =   8760
      TabIndex        =   5
      Top             =   7080
      Width           =   1815
   End
   Begin VB.CommandButton CommandClipBoard 
      Caption         =   "Copy To ClipBoard"
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
      Left            =   8760
      TabIndex        =   6
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton CommandZAFOptions 
      Caption         =   "ZAF Options"
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
      Left            =   8880
      TabIndex        =   11
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CheckBox CheckAllOptions 
      Caption         =   "Plot All Options"
      Height          =   255
      Left            =   8760
      TabIndex        =   17
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton CommandMACs 
      Caption         =   "MACs"
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
      Left            =   8880
      TabIndex        =   15
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CheckBox CheckAllMacs 
      Caption         =   "Plot All MACs"
      Height          =   255
      Left            =   8760
      TabIndex        =   16
      Top             =   5160
      Width           =   1695
   End
   Begin VB.OptionButton OptionBenceAlbee 
      Caption         =   "Option1"
      Height          =   615
      Index           =   2
      Left            =   8760
      TabIndex        =   14
      Top             =   2880
      Width           =   1815
   End
   Begin VB.OptionButton OptionBenceAlbee 
      Caption         =   "Option1"
      Height          =   615
      Index           =   1
      Left            =   8760
      TabIndex        =   13
      Top             =   2160
      Width           =   1815
   End
   Begin VB.OptionButton OptionBenceAlbee 
      Caption         =   "Option1"
      Height          =   615
      Index           =   0
      Left            =   8760
      TabIndex        =   12
      Top             =   1440
      Width           =   1815
   End
   Begin VB.ComboBox ComboPlotAlpha 
      Height          =   315
      Left            =   8880
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   960
      Width           =   1575
   End
   Begin VB.CheckBox CheckUseGridlines 
      Caption         =   "Use Grid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   7
      Top             =   6360
      Width           =   1095
   End
   Begin GraphsLib.Graph Graph1 
      Height          =   7335
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   8415
      _Version        =   327680
      _ExtentX        =   14843
      _ExtentY        =   12938
      _StockProps     =   96
      BorderStyle     =   1
      Background      =   "15~-1~-1~-1~-1~-1~-1"
      GraphType       =   0
      SymbolData      =   "13"
      SymbolSize      =   75
      ThickLines      =   0
      Toolbar         =   3
      LabelYFormat    =   ""
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   495
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label LabelMatrixCorrection 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   7560
      Width           =   8415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Binary Alpha"
      Height          =   255
      Left            =   8880
      TabIndex        =   9
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label LabelYPos 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   8640
      TabIndex        =   2
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label LabelXPos 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   8640
      TabIndex        =   1
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label LabelStdDev 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9600
      TabIndex        =   19
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "AvgDev%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9600
      TabIndex        =   20
      Top             =   5880
      Width           =   1095
   End
End
Attribute VB_Name = "FormPlotAlpha_GS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2015 by John J. Donovan
Option Explicit

Private Sub CheckAllMacs_Click()
If Not DebugMode Then On Error Resume Next
If FormPlotAlpha_GS.CheckAllMacs.Value = vbChecked Then
FormPlotAlpha_GS.CheckAllOptions.Enabled = False
FormPlotAlpha_GS.CommandMACs.Enabled = False
Else
FormPlotAlpha_GS.CommandMACs.Enabled = True
FormPlotAlpha_GS.CheckAllOptions.Enabled = True
End If
Call CalcZAFPlotAlphaFactors_GS
If ierror Then Exit Sub
End Sub

Private Sub CheckAllOptions_Click()
If Not DebugMode Then On Error Resume Next
If FormPlotAlpha_GS.CheckAllOptions.Value = vbChecked Then
FormPlotAlpha_GS.CheckAllMacs.Enabled = False
FormPlotAlpha_GS.CommandZAFOptions.Enabled = False
Else
FormPlotAlpha_GS.CommandZAFOptions.Enabled = True
FormPlotAlpha_GS.CheckAllMacs.Enabled = True
End If
Call CalcZAFPlotAlphaFactors_GS
If ierror Then Exit Sub
End Sub

Private Sub CheckUseGridlines_Click()
If Not DebugMode Then On Error Resume Next
If FormPlotAlpha_GS.CheckUseGridlines.Value = vbChecked Then
FormPlotAlpha_GS.Graph1.GridStyle = graphHorizontal_Vertical
FormPlotAlpha_GS.Graph1.GridLineStyle = graphDotted
Else
FormPlotAlpha_GS.Graph1.GridStyle = 0
End If
FormPlotAlpha_GS.Graph1.DrawMode = 2
End Sub

Private Sub ComboPlotAlpha_Click()
If Not DebugMode Then On Error Resume Next
Call CalcZAFPlotAlphaFactors_GS
End Sub

Private Sub Command1_Click()
If Not DebugMode Then On Error Resume Next
If CorrectionFlag% = 4 Then CorrectionFlag% = 3     ' force to polynomial alpha factors for now
Unload FormPlotAlpha_GS
End Sub

Private Sub Command2_Click()
' Zoom full
If Not DebugMode Then On Error Resume Next
' Change to variable origin
FormPlotAlpha_GS.Graph1.XAxisStyle = 1
FormPlotAlpha_GS.Graph1.YAxisStyle = 1
' Redraw graph
FormPlotAlpha_GS.Graph1.MousePointer = 0
FormPlotAlpha_GS.Graph1.DrawMode = 2
End Sub

Private Sub CommandClipBoard_Click()
If Not DebugMode Then On Error Resume Next
FormPlotAlpha_GS.Graph1.DrawMode = 3     ' blit to put in bitmap format (otherwise metafile format does not work correctly)
DoEvents
FormPlotAlpha_GS.Graph1.DrawMode = 4     ' copy to clipboard
End Sub

Private Sub CommandMACs_Click()
If Not DebugMode Then On Error Resume Next
Call GetZAFAllLoadMAC
If ierror Then Exit Sub
FormMAC.Show vbModal
Call CalcZAFPlotAlphaFactors_GS
If ierror Then Exit Sub
End Sub

Private Sub CommandPrint_Click()
If Not DebugMode Then On Error Resume Next
Call ZoomPrintGraph(FormPlotAlpha_GS.Graph1)
If ierror Then Exit Sub
End Sub

Private Sub CommandSaveData_Click()
If Not DebugMode Then On Error Resume Next
Call CalcZAFAlphaExportData_GS(FormPlotAlpha_GS)
If ierror Then Exit Sub
End Sub

Private Sub CommandZAFOptions_Click()
If Not DebugMode Then On Error Resume Next
Call GetZAFLoad
If ierror Then Exit Sub
FormGETZAF.Show vbModal
Call CalcZAFPlotAlphaFactors_GS
If ierror Then Exit Sub
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormPlotAlpha_GS)
HelpContextID = IOGetHelpContextID("FormPlotAlpha_GS")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub Graph1_SDKPaint()
If Not DebugMode Then On Error Resume Next
Call CalcZAFPlotAlphaFit(Int(0))    ' plot using Graphics Server code
If ierror Then Exit Sub
End Sub

Private Sub Graph1_SDKPress(PressStatus As Integer, PressX As Double, PressY As Double, PressDataX As Double, PressDataY As Double)
If Not DebugMode Then On Error Resume Next
Call ZoomSDKPress(PressStatus%, PressX#, PressY#, PressDataX#, PressDataY#, Int(1), FormPlotAlpha_GS)
If ierror Then Exit Sub
End Sub

Private Sub Graph1_SDKTrack(TrackX As Double, TrackY As Double, TrackDataX As Double, TrackDataY As Double)
If Not DebugMode Then On Error Resume Next

Dim astring As String
Dim tempx As Single, tempy As Single

' Tracking, load current graph data
tempx! = TrackDataX#
tempy! = TrackDataY#

' Format spectrometer position
astring$ = MiscAutoFormat$(tempx!)
FormPlotAlpha_GS.LabelXPos.Caption = astring$

' Format counts
astring$ = MiscAutoFormat$(tempy!)
FormPlotAlpha_GS.LabelYPos.Caption = astring$

' Zoom rectangle
Call ZoomSDKTrack(TrackX#, TrackY#)
If ierror Then Exit Sub

End Sub

Private Sub OptionBenceAlbee_Click(Index As Integer)
If Not DebugMode Then On Error Resume Next
CorrectionFlag% = Index% + 1
Call CalcZAFPlotAlphaFactors_GS
If ierror Then Exit Sub
End Sub
