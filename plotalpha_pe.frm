VERSION 5.00
Object = "{6E5043E8-C452-4A6A-B011-9B5687112610}#1.0#0"; "Pesgo32f.ocx"
Begin VB.Form FormPlotAlpha_PE 
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
   Begin Pesgo32fLib.Pesgo Pesgo1 
      Height          =   7335
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   8415
      _Version        =   65536
      _ExtentX        =   14843
      _ExtentY        =   12938
      _StockProps     =   96
      _AllProps       =   "PlotAlpha_PE.frx":0000
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
      TabIndex        =   17
      Top             =   7440
      Width           =   1815
   End
   Begin VB.CommandButton CommandZoomFull 
      BackColor       =   &H0080FFFF&
      Caption         =   "Zoom Full"
      Height          =   375
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7920
      Width           =   2055
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
      TabIndex        =   4
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
      TabIndex        =   5
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
      TabIndex        =   10
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CheckBox CheckAllOptions 
      Caption         =   "Plot All Options"
      Height          =   255
      Left            =   8760
      TabIndex        =   16
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
      TabIndex        =   14
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CheckBox CheckAllMacs 
      Caption         =   "Plot All MACs"
      Height          =   255
      Left            =   8760
      TabIndex        =   15
      Top             =   5160
      Width           =   1695
   End
   Begin VB.OptionButton OptionBenceAlbee 
      Caption         =   "Option1"
      Height          =   615
      Index           =   2
      Left            =   8760
      TabIndex        =   13
      Top             =   2880
      Width           =   1815
   End
   Begin VB.OptionButton OptionBenceAlbee 
      Caption         =   "Option1"
      Height          =   615
      Index           =   1
      Left            =   8760
      TabIndex        =   12
      Top             =   2160
      Width           =   1815
   End
   Begin VB.OptionButton OptionBenceAlbee 
      Caption         =   "Option1"
      Height          =   615
      Index           =   0
      Left            =   8760
      TabIndex        =   11
      Top             =   1440
      Width           =   1815
   End
   Begin VB.ComboBox ComboPlotAlpha 
      Height          =   315
      Left            =   8880
      Style           =   2  'Dropdown List
      TabIndex        =   7
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
      Left            =   9120
      TabIndex        =   6
      Top             =   6480
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CommandButton CommandClose 
      BackColor       =   &H00C0FFC0&
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   495
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label LabelMatrixCorrection 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   7560
      Width           =   8415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Binary Alpha"
      Height          =   255
      Left            =   8880
      TabIndex        =   8
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
Attribute VB_Name = "FormPlotAlpha_PE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2016 by John J. Donovan
Option Explicit

Private Sub CheckAllMacs_Click()
If Not DebugMode Then On Error Resume Next
If FormPlotAlpha_PE.CheckAllMacs.Value = vbChecked Then
FormPlotAlpha_PE.CheckAllOptions.Enabled = False
FormPlotAlpha_PE.CommandMACs.Enabled = False
Else
FormPlotAlpha_PE.CommandMACs.Enabled = True
FormPlotAlpha_PE.CheckAllOptions.Enabled = True
End If
Call CalcZAFPlotAlphaFactors_PE
If ierror Then Exit Sub
End Sub

Private Sub CheckAllOptions_Click()
If Not DebugMode Then On Error Resume Next
If FormPlotAlpha_PE.CheckAllOptions.Value = vbChecked Then
FormPlotAlpha_PE.CheckAllMacs.Enabled = False
FormPlotAlpha_PE.CommandZAFOptions.Enabled = False
Else
FormPlotAlpha_PE.CommandZAFOptions.Enabled = True
FormPlotAlpha_PE.CheckAllMacs.Enabled = True
End If
Call CalcZAFPlotAlphaFactors_PE
If ierror Then Exit Sub
End Sub

Private Sub CheckUseGridlines_Click()
If Not DebugMode Then On Error Resume Next
If FormPlotAlpha_PE.CheckUseGridlines.Value = vbChecked Then
FormPlotAlpha_PE.Pesgo1.GridLineControl = PEGLC_BOTH&          ' x and y grid
FormPlotAlpha_PE.Pesgo1.GridBands = True                       ' adds colour banding on background
Else
FormPlotAlpha_PE.Pesgo1.GridLineControl = PEGLC_NONE&
FormPlotAlpha_PE.Pesgo1.GridBands = False                      ' removes colour banding on background
End If
End Sub

Private Sub ComboPlotAlpha_Click()
If Not DebugMode Then On Error Resume Next
Call CalcZAFPlotAlphaFactors_PE
If ierror Then Exit Sub
End Sub

Private Sub CommandClipboard_Click()
If Not DebugMode Then On Error Resume Next
FormPlotAlpha_PE.Pesgo1.AllowExporting = True
FormPlotAlpha_PE.Pesgo1.ExportImageLargeFont = False
FormPlotAlpha_PE.Pesgo1.ExportImageDpi = 400
Call FormPlotAlpha_PE.Pesgo1.PEcopybitmaptoclipboard(FormPlotAlpha_PE.Pesgo1.Width / 10, FormPlotAlpha_PE.Pesgo1.Height / 10)
End Sub

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
Unload FormPlotAlpha_PE
End Sub

Private Sub CommandMACs_Click()
If Not DebugMode Then On Error Resume Next
Call GetZAFAllLoadMAC
If ierror Then Exit Sub
FormMAC.Show vbModal
Call CalcZAFPlotAlphaFactors_PE
If ierror Then Exit Sub
End Sub

Private Sub CommandPrint_Click()
If Not DebugMode Then On Error Resume Next
Call MiscPlotPrintGraph_PE(FormPlotAlpha_PE.Pesgo1)
If ierror Then Exit Sub
End Sub

Private Sub CommandSaveData_Click()
If Not DebugMode Then On Error Resume Next
Call CalcZAFAlphaExportData_PE(FormPlotAlpha_PE)
If ierror Then Exit Sub
End Sub

Private Sub CommandZAFOptions_Click()
If Not DebugMode Then On Error Resume Next
Call GetZAFLoad
If ierror Then Exit Sub
FormGETZAF.Show vbModal
Call CalcZAFPlotAlphaFactors_PE
If ierror Then Exit Sub
End Sub

Private Sub CommandZoomFull_Click()
If Not DebugMode Then On Error Resume Next
FormPlotAlpha_PE.Pesgo1.PEactions = UNDO_ZOOM&
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormPlotAlpha_PE)
HelpContextID = IOGetHelpContextID("FormPlotAlpha_PE")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub OptionBenceAlbee_Click(Index As Integer)
If Not DebugMode Then On Error Resume Next
CorrectionFlag% = Index% + 1
Call CalcZAFPlotAlphaFactors_PE
If ierror Then Exit Sub
End Sub

Private Sub Pesgo1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not DebugMode Then On Error Resume Next
Dim fX As Double, fY As Double      ' mouse position in graph coordinates

' Get mouse position in data units
Call MiscPlotTrack(Int(1), X!, Y!, fX#, fY#, FormPlotAlpha_PE.Pesgo1)
If ierror Then Exit Sub
   
' Format graph mouse position
If fX# <> 0# And fY# <> 0# Then
   FormPlotAlpha_PE.LabelXPos.Caption = MiscAutoFormat$(CSng(fX#))
   FormPlotAlpha_PE.LabelYPos.Caption = MiscAutoFormat$(CSng(fY#))
Else
   FormPlotAlpha_PE.LabelXPos.Caption = vbNullString
   FormPlotAlpha_PE.LabelYPos.Caption = vbNullString
End If
End Sub

Private Sub Pesgo1_ZoomOut()
If Not DebugMode Then On Error Resume Next
FormPlotAlpha_PE.Pesgo1.ManualScaleControlX = PEMSC_NONE&        ' automatically control X Axis
FormPlotAlpha_PE.Pesgo1.ManualScaleControlY = PEMSC_NONE&        ' automatically control Y Axis
End Sub
