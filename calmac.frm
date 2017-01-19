VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form FormMAIN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calmac (Calculate Mass Absorption Coefficients Using ABSORB.BAS)"
   ClientHeight    =   5055
   ClientLeft      =   1545
   ClientTop       =   2745
   ClientWidth     =   9585
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
   Icon            =   "CALMAC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5055
   ScaleWidth      =   9585
   Begin RichTextLib.RichTextBox TextLog 
      Height          =   1935
      Left            =   0
      TabIndex        =   10
      Top             =   1680
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   3413
      _Version        =   393217
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"CALMAC.frx":6E7FA
   End
   Begin VB.CommandButton CommandCalculateRange 
      Caption         =   "Calculate Range"
      Height          =   255
      Left            =   7680
      TabIndex        =   9
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox TextKeV 
      Height          =   285
      Left            =   6000
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.Timer TimerLogWindow 
      Left            =   720
      Top             =   3720
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   120
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.CommandButton CommandCalculateMAC 
      BackColor       =   &H0080FFFF&
      Caption         =   "Calculate MAC"
      Height          =   255
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   1815
   End
   Begin VB.ComboBox ComboAbsorber 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4080
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.ComboBox ComboElement 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.ComboBox ComboXRay 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   735
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   840
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   1296
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      ScrollBars      =   0
      Appearance      =   0
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      Caption         =   "KeV"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6000
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Absorber"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Element"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "X-Ray Line"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.Menu menuFile 
      Caption         =   "&File"
      Begin VB.Menu menuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu menuMethod 
      Caption         =   "&Method"
      Begin VB.Menu menuMethodMcMasterMACs 
         Caption         =   "McMaster MACs"
      End
      Begin VB.Menu menuMethodMAC30MACs 
         Caption         =   "MAC30 MACs"
      End
      Begin VB.Menu menuMethodJTAMACs 
         Caption         =   "MACJTA MACs"
      End
   End
   Begin VB.Menu menuOutput 
      Caption         =   "&Output"
      Begin VB.Menu menuOutputLogWindow 
         Caption         =   "Log Window Font"
      End
      Begin VB.Menu menuOutputDebugMode 
         Caption         =   "Debug Mode"
      End
      Begin VB.Menu menuOutputExtendedFormat 
         Caption         =   "Extended Format"
      End
      Begin VB.Menu menuOutputSaveToDiskLog 
         Caption         =   "Save To Disk Log"
      End
      Begin VB.Menu menuOutputViewDiskLog 
         Caption         =   "View Disk Log"
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "&Help"
      Begin VB.Menu menuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "FormMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2017 by John J. Donovan
Option Explicit

Private Sub ComboElement_Change()
If Not DebugMode Then On Error Resume Next
Call CalMacChange
If ierror Then Exit Sub
End Sub

Private Sub ComboElement_Click()
If Not DebugMode Then On Error Resume Next
Call CalMacChange
If ierror Then Exit Sub
End Sub

Private Sub CommandCalculateMAC_Click()
If Not DebugMode Then On Error Resume Next
Call CalMacCalculate
If ierror Then Exit Sub
End Sub

Private Sub CommandCalculateRange_Click()
If Not DebugMode Then On Error Resume Next
Screen.MousePointer = vbHourglass
Call CalMacCalculateRange
Screen.MousePointer = vbDefault
If ierror Then Exit Sub
End Sub

Private Sub Form_Activate()
If Not DebugMode Then On Error Resume Next

Static initialized As Integer

' Initialize the global variables
If initialized = False Then

' Initialize the files
Call InitFiles
If ierror Then End

' Load the PROBEWIN.INI file
Call InitINI
If ierror Then End

' Initialize arrays
Call InitData
If ierror Then End

Call CalMacLoad
If ierror Then Exit Sub

' Set default to use McMaster
FormMAIN.menuMethodMcMasterMACs.Checked = True

initialized = True
End If

End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next

' Set timer interval for write to Log Window
FormMAIN.TimerLogWindow.Interval = 500

' Load form and application icon
Call MiscLoadIcon(FormMAIN)

' Check if program is already running
If app.PrevInstance Then
msg$ = "CalMAC is already running, click OK, then type <ctrl> <esc> for the Task Manager and select CalMAC.EXE from the Task List"
MsgBox msg$, vbOKOnly + vbExclamation, "CalMAC"
ierror = True
End
End If

Call InitWindow(Int(2), MDBUserName$, Me)

' Help file
FormMAIN.HelpContextID = IOGetHelpContextID("FormMAIN")

End Sub

Private Sub Form_Resize()
If Not DebugMode Then On Error Resume Next

Dim i As Integer, temp As Integer

' Make text box (Log Window) full size of window
FormMAIN.TextLog.Left = 0
FormMAIN.TextLog.Width = FormMAIN.ScaleWidth
temp% = FormMAIN.ScaleHeight - FormMAIN.TextLog.Top
If temp% > 0 Then FormMAIN.TextLog.Height = temp%

' Make grid full size of window
FormMAIN.MSFlexGrid1.Left = 0
FormMAIN.MSFlexGrid1.Width = FormMAIN.ScaleWidth

For i% = 0 To FormMAIN.MSFlexGrid1.cols - 1
FormMAIN.MSFlexGrid1.ColWidth(i%) = FormMAIN.MSFlexGrid1.Width / FormMAIN.MSFlexGrid1.cols
Next i%

End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
End
End Sub

Private Sub menuExit_Click()
If Not DebugMode Then On Error Resume Next
Unload FormMAIN
End Sub

Private Sub menuHelpAbout_Click()
If Not DebugMode Then On Error Resume Next
FormABOUT.Show vbModal
End Sub

Private Sub menuMethodJTAMACs_Click()
If Not DebugMode Then On Error Resume Next
FormMAIN.menuMethodMcMasterMACs.Checked = False
FormMAIN.menuMethodMAC30MACs.Checked = False
FormMAIN.menuMethodJTAMACs.Checked = True
MACMode% = 2
End Sub

Private Sub menuMethodMAC30MACs_Click()
If Not DebugMode Then On Error Resume Next
FormMAIN.menuMethodMcMasterMACs.Checked = False
FormMAIN.menuMethodMAC30MACs.Checked = True
FormMAIN.menuMethodJTAMACs.Checked = False
MACMode% = 1
End Sub

Private Sub menuMethodMcMasterMACs_Click()
If Not DebugMode Then On Error Resume Next
FormMAIN.menuMethodMcMasterMACs.Checked = True
FormMAIN.menuMethodMAC30MACs.Checked = False
FormMAIN.menuMethodJTAMACs.Checked = False
MACMode% = 0
End Sub

Private Sub menuOutputDebugMode_Click()
If Not DebugMode Then On Error Resume Next
FormMAIN.menuOutputDebugMode.Checked = Not FormMAIN.menuOutputDebugMode.Checked
DebugMode = FormMAIN.menuOutputDebugMode.Checked
End Sub

Private Sub menuOutputExtendedFormat_Click()
If Not DebugMode Then On Error Resume Next
FormMAIN.menuOutputExtendedFormat.Checked = Not FormMAIN.menuOutputExtendedFormat.Checked
ExtendedFormat = FormMAIN.menuOutputExtendedFormat.Checked
End Sub

Private Sub menuOutputLogWindow_Click()
If Not DebugMode Then On Error Resume Next
Call IOLogFont
End Sub

Private Sub menuOutputSaveToDiskLog_Click()
' This routine toggles the "SaveToDisk" flag and opens or closes the .OUT file
If Not DebugMode Then On Error Resume Next
' Perform file operations and update flag
If Not SaveToDisk Then
Call IOOpenOUTFile(FormMAIN)
Else
Call IOCloseOUTFile
End If
End Sub

Private Sub menuOutputViewDiskLog_Click()
If Not DebugMode Then On Error Resume Next
Call IOViewLog
If ierror Then Exit Sub
End Sub

Private Sub TextKeV_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextLog_KeyPress(KeyAscii As Integer)
If Not DebugMode Then On Error Resume Next
Call IOSendLog(KeyAscii%)
If ierror Then Exit Sub
End Sub

Private Sub TimerLogWindow_Timer()
If Not DebugMode Then On Error Resume Next
Call IODumpLog
End Sub

