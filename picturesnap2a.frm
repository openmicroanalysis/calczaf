VERSION 5.00
Begin VB.Form FormPICTURESNAP2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Picture Snap Calibration"
   ClientHeight    =   11790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11790
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TextScan 
      Height          =   285
      Left            =   3960
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   11400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox TextkeV 
      Height          =   285
      Left            =   3480
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   11040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox TextMag 
      Height          =   285
      Left            =   4440
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   11040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton CommandDisplayCalibrationPoints 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Display Calibration Points"
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
      TabIndex        =   36
      ToolTipText     =   "Click this button to have the program display the location of the calibration points on the calibrated image"
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   "Point #3 Calibration"
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
      Height          =   3615
      Left            =   120
      TabIndex        =   30
      Top             =   8040
      Width           =   3135
      Begin VB.TextBox TextZStage3 
         Height          =   285
         Left            =   2040
         TabIndex        =   40
         ToolTipText     =   "Enter the Y stage coordinate of the first calibration point"
         Top             =   2760
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox TextYStage3 
         Height          =   285
         Left            =   2040
         TabIndex        =   11
         ToolTipText     =   "Enter the Y stage coordinate of the third calibration point"
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox TextXStage3 
         Height          =   285
         Left            =   2040
         TabIndex        =   10
         ToolTipText     =   "Enter the X stage coordinate of the third calibration point"
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton CommandPickPixelCoordinate3 
         Caption         =   "Pick Pixel Coordinate on Picture"
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
         Left            =   360
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "Click this button and then click the picture image using the mouse to load the pixel coordinate of the third calibration point"
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox TextYPixel3 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "Click the image to read the Y pixel coordinate of the third calibration point"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox TextXPixel3 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   8
         ToolTipText     =   "Click the image to read the X pixel coordinate of the third calibration point"
         Top             =   480
         Width           =   975
      End
      Begin VB.Label LabelZStage3 
         Caption         =   "Z Stage Coordinate"
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
         TabIndex        =   44
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label14 
         Caption         =   "Y Stage Coordinate"
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
         TabIndex        =   35
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label13 
         Caption         =   "X Stage Coordinate"
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
         TabIndex        =   34
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "Y Pixel Coordinate"
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
         TabIndex        =   33
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "X Pixel Coordinate"
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
         TabIndex        =   32
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.OptionButton OptionPictureSnapMode 
      Caption         =   "Three Points"
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
      Left            =   3480
      TabIndex        =   29
      ToolTipText     =   $"PictureSnap2a.frx":0000
      Top             =   1080
      Width           =   1575
   End
   Begin VB.OptionButton OptionPictureSnapMode 
      Caption         =   "Two Points"
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
      Left            =   3480
      TabIndex        =   28
      ToolTipText     =   "Use this option for calibrating rectangular samples (e.g., petrographic thin sections) that are orthogonal to the stage"
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton CommandCalibratePicture 
      BackColor       =   &H0000FFFF&
      Caption         =   "Calibrate Picture"
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Click this button when both calibration coordinates in pixel and stage coordinates have been entered to calibrate the picture"
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Point #2 Calibration"
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
      Height          =   3615
      Left            =   120
      TabIndex        =   19
      Top             =   4080
      Width           =   3135
      Begin VB.TextBox TextZStage2 
         Height          =   285
         Left            =   2040
         TabIndex        =   39
         ToolTipText     =   "Enter the Y stage coordinate of the first calibration point"
         Top             =   2760
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox TextXPixel2 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   4
         ToolTipText     =   "Click the image to read the X pixel coordinate of the second calibration point"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox TextYPixel2 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "Click the image to read the Y pixel coordinate of the second calibration point"
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton CommandPickPixelCoordinate2 
         Caption         =   "Pick Pixel Coordinate on Picture"
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
         Left            =   360
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Click this button and then click the picture image using the mouse to load the pixel coordinate of the second calibration point"
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox TextXStage2 
         Height          =   285
         Left            =   2040
         TabIndex        =   6
         ToolTipText     =   "Enter the X stage coordinate of the second calibration point"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox TextYStage2 
         Height          =   285
         Left            =   2040
         TabIndex        =   7
         ToolTipText     =   "Enter the Y stage coordinate of the second calibration point"
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label LabelZStage2 
         Caption         =   "Z Stage Coordinate"
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
         TabIndex        =   43
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "X Pixel Coordinate"
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
         TabIndex        =   24
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Y Pixel Coordinate"
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
         TabIndex        =   23
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "X Stage Coordinate"
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
         TabIndex        =   22
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Y Stage Coordinate"
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
         TabIndex        =   21
         Top             =   2400
         Width           =   1815
      End
   End
   Begin VB.CommandButton CommandClose 
      BackColor       =   &H00008000&
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   120
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Point #1 Calibration"
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
      Height          =   3615
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   3135
      Begin VB.TextBox TextZStage1 
         Height          =   285
         Left            =   2040
         TabIndex        =   38
         ToolTipText     =   "Enter the Y stage coordinate of the first calibration point"
         Top             =   2760
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox TextYStage1 
         Height          =   285
         Left            =   2040
         TabIndex        =   3
         ToolTipText     =   "Enter the Y stage coordinate of the first calibration point"
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox TextXStage1 
         Height          =   285
         Left            =   2040
         TabIndex        =   2
         ToolTipText     =   "Enter the X stage coordinate of the first calibration point"
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton CommandPickPixelCoordinate1 
         Caption         =   "Pick Pixel Coordinate on Picture"
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
         Left            =   360
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Click this button and then click the picture image using the mouse to load the pixel coordinate of the first calibration point"
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox TextYPixel1 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "Click the image to read the Y pixel coordinate of the first calibration point"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox TextXPixel1 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   0
         ToolTipText     =   "Click the image to read the X pixel coordinate of the first calibration point"
         Top             =   480
         Width           =   975
      End
      Begin VB.Label LabelZStage1 
         Caption         =   "Z Stage Coordinate"
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
         TabIndex        =   42
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Y Stage Coordinate"
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
         TabIndex        =   18
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "X Stage Coordinate"
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
         TabIndex        =   17
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Y Pixel Coordinate"
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
         TabIndex        =   15
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "X Pixel Coordinate"
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
         TabIndex        =   14
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Label LabelCalibration 
      Alignment       =   2  'Center
      Caption         =   "Image Is Calibrated!"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3480
      TabIndex        =   41
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "When using a three point calibration, the program automatically includes a Z correction for stage sample tilt!"
      Height          =   1095
      Left            =   3360
      TabIndex        =   37
      Top             =   8160
      Width           =   1935
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "X and Y Pixel Coordinates are actually given in ""Twip"" units! (1440 twips per logical inch)"
      Height          =   855
      Left            =   3360
      TabIndex        =   27
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   $"PictureSnap2a.frx":008C
      Height          =   2055
      Left            =   3360
      TabIndex        =   26
      Top             =   1440
      Width           =   1935
   End
End
Attribute VB_Name = "FormPICTURESNAP2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2016 by John J. Donovan
Option Explicit

Private Sub CommandCalibratePicture_Click()
If Not DebugMode Then On Error Resume Next
Call PictureSnapCalibrate(Int(1))
If ierror Then Exit Sub
End Sub

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
Unload FormPICTURESNAP2
End Sub

Private Sub CommandDisplayCalibrationPoints_Click()
If Not DebugMode Then On Error Resume Next
PictureSnapDisplayCalibrationPointsFlag = Not PictureSnapDisplayCalibrationPointsFlag
If PictureSnapDisplayCalibrationPointsFlag Then
FormPICTURESNAP2.CommandDisplayCalibrationPoints.Caption = "Do Not Display Calibration Points"
Else
FormPICTURESNAP2.CommandDisplayCalibrationPoints.Caption = "Display Calibration Points"
End If
FormPICTURESNAP.Picture2.Refresh
If FormPICTURESNAP3.Visible Then FormPICTURESNAP3.Image1.Refresh
End Sub

Private Sub CommandPickPixelCoordinate1_Click()
If Not DebugMode Then On Error Resume Next
Call PictureSnapCalibratePoint(Int(1))
If ierror Then Exit Sub
End Sub

Private Sub CommandPickPixelCoordinate2_Click()
If Not DebugMode Then On Error Resume Next
Call PictureSnapCalibratePoint(Int(2))
If ierror Then Exit Sub
End Sub

Private Sub CommandPickPixelCoordinate3_Click()
If Not DebugMode Then On Error Resume Next
Call PictureSnapCalibratePoint(Int(3))
If ierror Then Exit Sub
End Sub

Private Sub Form_Activate()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)   ' save here also because in case user moves window and activate events occurs before closing window
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call MiscAlwaysOnTop(True, FormPICTURESNAP2)
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormPICTURESNAP2)
HelpContextID = IOGetHelpContextID("FormPICTURESNAP2")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
If FormPICTURESELECT.Visible Then Unload FormPICTURESELECT
Screen.MousePointer = vbDefault
End Sub

Private Sub OptionPictureSnapMode_Click(Index As Integer)
If Not DebugMode Then On Error Resume Next
Call PictureSnapSaveMode(Index%)
If ierror Then Exit Sub
End Sub

Private Sub TextXPixel1_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextXPixel2_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextXPixel3_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextXStage1_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextXStage2_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextXStage3_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextYPixel1_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextYPixel2_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextYPixel3_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextYStage1_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextYStage2_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextYStage3_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextZStage1_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextZStage2_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextZStage3_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub
