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
   Begin VB.Timer TimerPictureSnap 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8280
      Top             =   0
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   5280
      Width           =   7095
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5295
      Left            =   7080
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   5295
      Left            =   0
      ScaleHeight     =   5235
      ScaleWidth      =   7035
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   1095
         Left            =   120
         ScaleHeight     =   1035
         ScaleWidth      =   1035
         TabIndex        =   4
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         Height          =   1575
         Left            =   2520
         ScaleHeight     =   1515
         ScaleWidth      =   1515
         TabIndex        =   1
         Top             =   1920
         Width           =   1575
      End
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
' (c) Copyright 1995-2015 by John J. Donovan
Option Explicit

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
Dim xdim As Single, ydim As Single

' Move the Container PicBox to fill the screen but leave room for the scrollbars
If Me.ScaleWidth = 0 Or Me.ScaleHeight = 0 Then Exit Sub
FormPICTURESNAP.Picture1.Move 0, 0, Me.ScaleWidth - FormPICTURESNAP.VScroll1.Width, Me.ScaleHeight - FormPICTURESNAP.HScroll1.Height

' Move the scrollbar to the far right and make it as high as the screen
FormPICTURESNAP.VScroll1.Move Me.ScaleWidth - FormPICTURESNAP.VScroll1.Width, 0, FormPICTURESNAP.VScroll1.Width, Me.ScaleHeight
' Move the scrollbar to the far bottom and make it as wide as the screen (minus the vertical scrollbar)
FormPICTURESNAP.HScroll1.Move 0, Me.ScaleHeight - FormPICTURESNAP.HScroll1.Height, Me.ScaleWidth - FormPICTURESNAP.VScroll1.Width, FormPICTURESNAP.HScroll1.Height

' Set the borderstyle for pic2 to no border
FormPICTURESNAP.Picture2.BorderStyle = 0

' Size pic2 so it is bigger that view window
If FormPICTURESNAP.Picture2.Picture.Type > 0 Then   ' bitmap
xdim! = FormPICTURESNAP.Picture2.ScaleX(FormPICTURESNAP.Picture2.Picture.Width, vbHimetric, vbTwips)
ydim! = FormPICTURESNAP.Picture2.ScaleY(FormPICTURESNAP.Picture2.Picture.Height, vbHimetric, vbTwips)
FormPICTURESNAP.Picture2.Move 0, 0, xdim!, ydim!
Else
FormPICTURESNAP.Picture2.Move 0, 0, FormPICTURESNAP.Picture1.ScaleWidth * 30, FormPICTURESNAP.Picture1.ScaleHeight * 20
End If
    
' Set the scroll bars change settings
If Abs(FormPICTURESNAP.Picture1.ScaleHeight - FormPICTURESNAP.Picture2.ScaleHeight) < MAXINTEGER% Then
FormPICTURESNAP.VScroll1.Max = Abs(FormPICTURESNAP.Picture1.ScaleHeight - FormPICTURESNAP.Picture2.ScaleHeight)
If FormPICTURESNAP.VScroll1.Max / 10 > 1 Then FormPICTURESNAP.VScroll1.LargeChange = FormPICTURESNAP.VScroll1.Max / 10
If FormPICTURESNAP.VScroll1.Max / 100 > 1 Then FormPICTURESNAP.VScroll1.SmallChange = FormPICTURESNAP.VScroll1.Max / 100
End If

If Abs(FormPICTURESNAP.Picture1.ScaleWidth - FormPICTURESNAP.Picture2.ScaleWidth) < MAXINTEGER% Then
FormPICTURESNAP.HScroll1.Max = Abs(FormPICTURESNAP.Picture1.ScaleWidth - FormPICTURESNAP.Picture2.ScaleWidth)
If FormPICTURESNAP.HScroll1.Max / 10 > 1 Then FormPICTURESNAP.HScroll1.LargeChange = FormPICTURESNAP.HScroll1.Max / 10
If FormPICTURESNAP.HScroll1.Max / 100 > 1 Then FormPICTURESNAP.HScroll1.SmallChange = FormPICTURESNAP.HScroll1.Max / 100
End If

' To help scale bar redraw
HScroll1_Scroll
VScroll1_Scroll
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
Call PictureSnapLoadPositions(Int(1))
If ierror Then Exit Sub
End Sub

Private Sub menuDisplayUnknowns_Click()
If Not DebugMode Then On Error Resume Next
FormPICTURESNAP.menuDisplayUnknowns.Checked = Not FormPICTURESNAP.menuDisplayUnknowns.Checked
FormPICTURESNAP.menuDisplayStandards.Checked = False
FormPICTURESNAP.menuDisplayWavescans.Checked = False
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
SavePicture FormPICTURESNAP.Picture2, tfilename$     ' does not save graphics methods
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

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Not DebugMode Then On Error Resume Next
BitMapButton% = Button%
BitMapX! = x!
BitMapY! = Y!   ' store for double-click and map calibrate
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Not DebugMode Then On Error Resume Next
Call PictureSnapUpdateCursor(Int(0), x!, Y!)
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

Private Sub VScroll1_Change()
If Not DebugMode Then On Error Resume Next
VScroll1_Scroll
End Sub

Private Sub VScroll1_Scroll()
If Not DebugMode Then On Error Resume Next
FormPICTURESNAP.Picture2.Top = -FormPICTURESNAP.VScroll1.value
Call PictureSnapResetScaleBar
End Sub

Private Sub HScroll1_Change()
If Not DebugMode Then On Error Resume Next
HScroll1_Scroll
End Sub

Private Sub HScroll1_Scroll()
If Not DebugMode Then On Error Resume Next
FormPICTURESNAP.Picture2.Left = -FormPICTURESNAP.HScroll1.value
Call PictureSnapResetScaleBar
End Sub


