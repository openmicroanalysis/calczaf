VERSION 5.00
Begin VB.Form FormPICTURESNAP 
   Caption         =   "Picture Snap"
   ClientHeight    =   6765
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3195
      ScaleWidth      =   3555
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.Timer TimerPictureSnap 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8280
      Top             =   0
   End
   Begin VB.Menu menuFile 
      Caption         =   "&File"
      Begin VB.Menu menuFileOpenImage 
         Caption         =   "Open Image File"
      End
      Begin VB.Menu menuFileImportPrbImg 
         Caption         =   "Import PrbImg File (Probe Image)"
      End
      Begin VB.Menu menuFileSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu menuFileImportGridFile 
         Caption         =   "Import Grid (.GRD) File As Image"
      End
      Begin VB.Menu menuFileSaveAsGRD 
         Caption         =   "Save Image As Grid (.GRD) File"
         Enabled         =   0   'False
      End
      Begin VB.Menu menuFileSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu menuFileClipboard1 
         Caption         =   "Copy To Clipboard (method 1)"
         Enabled         =   0   'False
      End
      Begin VB.Menu menuFileClipboard2 
         Caption         =   "Copy To Clipboard (method 2)"
         Enabled         =   0   'False
      End
      Begin VB.Menu menuFileSaveAsBMPOnly 
         Caption         =   "Save As BMP (no coordinate calibration)"
         Enabled         =   0   'False
      End
      Begin VB.Menu menuFileSaveAsBMP 
         Caption         =   "Save As BMP (with graphics objects)"
         Enabled         =   0   'False
      End
      Begin VB.Menu menuFileSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu menuFilePrintSetup 
         Caption         =   "Print Setup"
         Enabled         =   0   'False
      End
      Begin VB.Menu menuFilePrint 
         Caption         =   "Print"
         Enabled         =   0   'False
      End
      Begin VB.Menu menuFileSeparator4 
         Caption         =   "-"
      End
      Begin VB.Menu menuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu menuWindow 
      Caption         =   "&Window"
      Begin VB.Menu menuWindowCalibrate 
         Caption         =   "Calibrate Image To Stage Coordinates"
      End
      Begin VB.Menu menuWindowSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu menuWindowFullPicture 
         Caption         =   "Full Image Picture View"
      End
   End
   Begin VB.Menu menuDisplay 
      Caption         =   "&Display"
      Begin VB.Menu menuDisplayStandards 
         Caption         =   "Digitized Standard Position Samples"
      End
      Begin VB.Menu menuDisplayUnknowns 
         Caption         =   "Digitized Unknown Position Samples"
      End
      Begin VB.Menu menuDisplayWavescans 
         Caption         =   "Digitized Wavescan Position Samples"
      End
      Begin VB.Menu menuDisplaySeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu menuDisplayLongLabels 
         Caption         =   "Digitized Position Sample Long Labels (Sample and Line Numbers)"
      End
      Begin VB.Menu menuDisplayShortLabels 
         Caption         =   "Digitized Position Sample Short Labels (Line Numbers Only)"
      End
      Begin VB.Menu menuDisplaySeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu menuDisplayUseBlackScaleBar 
         Caption         =   "Use Black Scaler Bar"
      End
      Begin VB.Menu menuDisplayDisplayDigitizedPositionsForSelectedPositionSampleOnly 
         Caption         =   "Display Digitized Positions For Selected Position Sample Only"
      End
   End
   Begin VB.Menu menuMisc 
      Caption         =   "&Misc"
      Begin VB.Menu menuMiscUseBeamBlankForStageMotion 
         Caption         =   "Use Beam Blank For Stage Motion"
      End
      Begin VB.Menu menuMiscUseRightMouseClickToDigitize 
         Caption         =   "Use Right Mouse Click To Digitize Positions"
      End
      Begin VB.Menu menuMiscMaintainAspectRatioOfFullViewWindow 
         Caption         =   "Maintain Aspect Ratio of Full View Window"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "FormPICTURESNAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2016 by John J. Donovan
Option Explicit

' ImageViewer CP Pro SDK OCX version

Dim BitMapButton As Integer
Dim BitMapX As Single
Dim BitMapY As Single

Private Sub Form_Activate()
If Not DebugMode Then On Error Resume Next
' Activate timer only on form activate event!
FormPICTURESNAP.TimerPictureSnap.Enabled = True
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormPICTURESNAP)
HelpContextID = IOGetHelpContextID("FormPICTURESNAP")
End Sub

Private Sub Form_Resize()
If Not DebugMode Then On Error Resume Next


End Sub
 
Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call PictureSnapSave
If ierror Then Exit Sub
Call InitWindow(Int(1), MDBUserName$, Me)
Unload FormPICTURESNAP2    ' unload calibration form in case it is loaded
Unload FormPICTURESNAP3    ' unload full window view in case it is loaded
FormPICTURESNAP.TimerPictureSnap.Enabled = False
End Sub

Private Sub menuDisplayDisplayDigitizedPositionsForSelectedPositionSampleOnly_Click()
If Not DebugMode Then On Error Resume Next
FormPICTURESNAP.menuDisplayDisplayDigitizedPositionsForSelectedPositionSampleOnly.Checked = Not FormPICTURESNAP.menuDisplayDisplayDigitizedPositionsForSelectedPositionSampleOnly.Checked
FormPICTURESNAP.menuDisplayWavescans.Checked = False
FormPICTURESNAP.menuDisplayStandards.Checked = False
FormPICTURESNAP.menuDisplayUnknowns.Checked = False
Call PictureSnapLoadPositions(Int(0))
If ierror Then Exit Sub
End Sub

Private Sub menuDisplayLongLabels_Click()
If Not DebugMode Then On Error Resume Next
FormPICTURESNAP.menuDisplayLongLabels.Checked = Not FormPICTURESNAP.menuDisplayLongLabels.Checked
If FormPICTURESNAP.menuDisplayLongLabels.Checked Then FormPICTURESNAP.menuDisplayShortLabels.Checked = False
Call PictureSnapLoadPositions(Int(0))
If ierror Then Exit Sub
End Sub

Private Sub menuDisplayShortLabels_Click()
If Not DebugMode Then On Error Resume Next
FormPICTURESNAP.menuDisplayShortLabels.Checked = Not FormPICTURESNAP.menuDisplayShortLabels.Checked
If FormPICTURESNAP.menuDisplayLongLabels.Checked Then FormPICTURESNAP.menuDisplayLongLabels.Checked = False
Call PictureSnapLoadPositions(Int(0))
If ierror Then Exit Sub
End Sub

Private Sub menuDisplayStandards_Click()
If Not DebugMode Then On Error Resume Next
FormPICTURESNAP.menuDisplayStandards.Checked = Not FormPICTURESNAP.menuDisplayStandards.Checked
FormPICTURESNAP.menuDisplayUnknowns.Checked = False
FormPICTURESNAP.menuDisplayWavescans.Checked = False
FormPICTURESNAP.menuDisplayDisplayDigitizedPositionsForSelectedPositionSampleOnly.Checked = False
Call PictureSnapLoadPositions(Int(1))
If ierror Then Exit Sub
End Sub

Private Sub menuDisplayUnknowns_Click()
If Not DebugMode Then On Error Resume Next
FormPICTURESNAP.menuDisplayUnknowns.Checked = Not FormPICTURESNAP.menuDisplayUnknowns.Checked
FormPICTURESNAP.menuDisplayStandards.Checked = False
FormPICTURESNAP.menuDisplayWavescans.Checked = False
FormPICTURESNAP.menuDisplayDisplayDigitizedPositionsForSelectedPositionSampleOnly.Checked = False
Call PictureSnapLoadPositions(Int(2))
If ierror Then Exit Sub
End Sub

Private Sub menuDisplayUseBlackScaleBar_Click()
If Not DebugMode Then On Error Resume Next
FormPICTURESNAP.menuDisplayUseBlackScaleBar.Checked = Not FormPICTURESNAP.menuDisplayUseBlackScaleBar.Checked
End Sub

Private Sub menuDisplayWavescans_Click()
If Not DebugMode Then On Error Resume Next
FormPICTURESNAP.menuDisplayWavescans.Checked = Not FormPICTURESNAP.menuDisplayWavescans.Checked
FormPICTURESNAP.menuDisplayStandards.Checked = False
FormPICTURESNAP.menuDisplayUnknowns.Checked = False
FormPICTURESNAP.menuDisplayDisplayDigitizedPositionsForSelectedPositionSampleOnly.Checked = False
Call PictureSnapLoadPositions(Int(3))
If ierror Then Exit Sub
End Sub

Private Sub menuExit_Click()
If Not DebugMode Then On Error Resume Next
Unload FormPICTURESNAP
End Sub

Private Sub menuFileClipboard2_Click()
If Not DebugMode Then On Error Resume Next
Call PictureSnapPrintOrClipboard(Int(3), FormPICTURESNAP)
If ierror Then Exit Sub
End Sub

Private Sub menuFileClipboard1_Click()
If Not DebugMode Then On Error Resume Next
Call PictureSnapPrintOrClipboard(Int(2), FormPICTURESNAP)
If ierror Then Exit Sub
End Sub

Private Sub menuFileImportGridFile_Click()
If Not DebugMode Then On Error Resume Next
Call PictureSnapFileOpen(Int(2), vbNullString, FormPICTURESNAP)
If ierror Then Exit Sub
End Sub

Private Sub menuFileImportPrbImg_Click()
If Not DebugMode Then On Error Resume Next
Call PictureSnapImportPrbImg
If ierror Then Exit Sub
End Sub

Private Sub menuFileOpenImage_Click()
If Not DebugMode Then On Error Resume Next
Call PictureSnapFileOpen(Int(1), vbNullString, FormPICTURESNAP)
If ierror Then Exit Sub
End Sub

Private Sub menuFilePrint_Click()
If Not DebugMode Then On Error Resume Next
Call PictureSnapPrintOrClipboard(Int(1), FormPICTURESNAP)
If ierror Then Exit Sub
End Sub

Private Sub menuFilePrintSetup_Click()
If Not DebugMode Then On Error Resume Next
Call IOPrintSetup
If ierror Then Exit Sub
End Sub

Private Sub menuFileSaveAsBMP_Click()
If Not DebugMode Then On Error Resume Next
Call PictureSnapPrintOrClipboard(Int(4), FormPICTURESNAP)
If ierror Then Exit Sub
End Sub

Private Sub menuFileSaveAsBMPOnly_Click()
If Not DebugMode Then On Error Resume Next
' Will not be flipped properly if default polarity "config" is different than file polarity "config"
Dim tfilename As String
tfilename$ = MiscGetFileNameNoExtension$(ProbeDataFile$) & "_" & "PictureSnap" & ".bmp"
Call IOGetFileName(Int(1), "BMP", tfilename$, FormPICTURESNAP)
If ierror Then Exit Sub
'SavePicture FormPICTURESNAP.Picture2, tfilename$     ' does not save graphics methods
End Sub

Private Sub menuFileSaveAsGRD_Click()
If Not DebugMode Then On Error Resume Next
Call PictureSnapSaveGridFile
If ierror Then Exit Sub
End Sub

Private Sub menuMiscMaintainAspectRatioOfFullViewWindow_Click()
If Not DebugMode Then On Error Resume Next
FormPICTURESNAP.menuMiscMaintainAspectRatioOfFullViewWindow.Checked = Not FormPICTURESNAP.menuMiscMaintainAspectRatioOfFullViewWindow.Checked
End Sub

Private Sub menuMiscUseBeamBlankForStageMotion_Click()
If Not DebugMode Then On Error Resume Next
FormPICTURESNAP.menuMiscUseBeamBlankForStageMotion.Checked = Not FormPICTURESNAP.menuMiscUseBeamBlankForStageMotion.Checked
End Sub

Private Sub menuMiscUseRightMouseClickToDigitize_Click()
If Not DebugMode Then On Error Resume Next
FormPICTURESNAP.menuMiscUseRightMouseClickToDigitize.Checked = Not FormPICTURESNAP.menuMiscUseRightMouseClickToDigitize.Checked
UseRightMouseClickToDigitizeFlag = FormPICTURESNAP.menuMiscUseRightMouseClickToDigitize.Checked
If UseRightMouseClickToDigitizeFlag And Not FormAUTOMATE.Visible Then
msg$ = "Please selected a position sample in the Automate! window before starting to digitize stage positions."
MsgBox msg$, vbOKOnly + vbExclamation, "FORMPICTURESNAP"
Exit Sub
End If
End Sub

Private Sub menuWindowCalibrate_Click()
If Not DebugMode Then On Error Resume Next
Call PictureSnapCalibrateLoad(Int(1))
If ierror Then Exit Sub
End Sub

Private Sub menuWindowFullPicture_Click()
If Not DebugMode Then On Error Resume Next
Call PictureSnapLoadFullWindow
If ierror Then Exit Sub
End Sub

Private Sub Picture2_Click()
' Transfer to PictureSnapSelect
Call PictureSnapSelectUpdate(BitMapX!, BitMapY!)
If ierror Then Exit Sub
PictureSnapClicked = True
DoEvents

' Digitize right clicked position to position database (if menu is checked)
If BitMapButton% = vbRightButton And FormPICTURESNAP.menuMiscUseRightMouseClickToDigitize.Checked Then
Call PictureSnapDigitizePoint(BitMapX!, BitMapY!)
If ierror Then Exit Sub
End If
End Sub

Private Sub Picture2_DblClick()
If Not DebugMode Then On Error Resume Next
If BitMapButton% = vbLeftButton Then
Call PictureSnapStageMove(BitMapX!, BitMapY!)
If ierror Then Exit Sub
End If
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not DebugMode Then On Error Resume Next
BitMapButton% = Button%
BitMapX! = X!
BitMapY! = Y!   ' store for double-click and map calibrate
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not DebugMode Then On Error Resume Next
Call PictureSnapUpdateCursor(Int(0), X!, Y!)
If ierror Then Exit Sub
End Sub

Private Sub TimerPictureSnap_Timer()
If Not DebugMode Then On Error Resume Next
Call PictureSnapDrawCurrentPosition
If ierror Then Exit Sub
Call PictureSnapDisplayPositions
If ierror Then Exit Sub
Call PictureSnapDrawScaleBar
If ierror Then Exit Sub
End Sub

