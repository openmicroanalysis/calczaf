VERSION 5.00
Begin VB.Form FormPICTURESNAP2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PictureSnap Image Calibration"
   ClientHeight    =   12075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12075
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   39
      ToolTipText     =   "Click this button to have the program display the location of the calibration points on the calibrated image"
      Top             =   5880
      Width           =   2415
   End
   Begin VB.CommandButton CommandLoadACQCalibration 
      Caption         =   "Load ACQ File for Calibration"
      Height          =   495
      Left            =   3480
      TabIndex        =   57
      ToolTipText     =   "Select an existing ACQ file to calibrate the currently loaded image"
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Frame Frame4 
      Caption         =   "Light Mode"
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
      Height          =   855
      Left            =   3600
      TabIndex        =   47
      Top             =   7080
      Width           =   2175
      Begin VB.CommandButton CommandLightModeOff 
         Caption         =   "Off"
         Height          =   255
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Turn optical light off (reflected or transmitted)"
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton CommandLightModeOn 
         Caption         =   "On"
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Turn optical light on (reflected or transmitted)"
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton CommandLightModeTransmitted 
         Caption         =   "Tran"
         Height          =   255
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Select to switch optical mode to transmitted light"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CommandLightModeReflected 
         Caption         =   "Refl"
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Select to switch optical mode to reflected light"
         Top             =   240
         Width           =   735
      End
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
      Height          =   3735
      Left            =   120
      TabIndex        =   32
      Top             =   8280
      Width           =   3135
      Begin VB.CommandButton CommandMoveTo3 
         Caption         =   "Move To"
         Height          =   495
         Left            =   2280
         TabIndex        =   46
         ToolTipText     =   "Click this button to move to the stage coordinates for point #3 calibration"
         Top             =   3120
         Width           =   735
      End
      Begin VB.CommandButton CommandReadCurrentStageCoordinate3 
         Caption         =   "Read Current Stage Coordinate"
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
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "Click this button to load the current stage coordinates for the third calibration point"
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox TextZStage3 
         Height          =   285
         Left            =   2040
         TabIndex        =   43
         ToolTipText     =   "Enter the Z stage coordinate of the first calibration point"
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
         Left            =   240
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "Click this button and then click the picture image using the mouse to load the pixel coordinate of the third calibration point"
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox TextYPixel3 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "Click the image to read the Y pixel coordinate of the third calibration point"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox TextXPixel3 
         BackColor       =   &H80000004&
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
         TabIndex        =   55
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
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
      Left            =   3960
      TabIndex        =   31
      ToolTipText     =   $"PictureSnap2.frx":0000
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
      Left            =   3960
      TabIndex        =   30
      ToolTipText     =   "Use this option for calibrating rectangular samples (e.g., petrographic thin sections) that are orthogonal to the stage"
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton CommandCalibratePicture 
      BackColor       =   &H00C0FFFF&
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
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "Click this button when both calibration coordinates in pixel and stage coordinates have been entered to calibrate the picture"
      Top             =   3120
      Width           =   2415
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
      Height          =   3735
      Left            =   120
      TabIndex        =   20
      Top             =   4200
      Width           =   3135
      Begin VB.CommandButton CommandMoveTo2 
         Caption         =   "Move To"
         Height          =   495
         Left            =   2160
         TabIndex        =   45
         ToolTipText     =   "Click this button to move to the stage coordinates for point #2 calibration"
         Top             =   3120
         Width           =   735
      End
      Begin VB.CommandButton CommandReadCurrentStageCoordinate2 
         Caption         =   "Read Current Stage Coordinate"
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
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Click this button to load the current stage coordinates for the second calibration point"
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox TextZStage2 
         Height          =   285
         Left            =   2040
         TabIndex        =   42
         ToolTipText     =   "Enter the Z stage coordinate of the first calibration point"
         Top             =   2760
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox TextXPixel2 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   4
         ToolTipText     =   "Click the image to read the X pixel coordinate of the second calibration point"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox TextYPixel2 
         BackColor       =   &H80000004&
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
         Left            =   240
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Click this button and then click the picture image using the mouse to load the pixel coordinate of the second calibration point"
         Top             =   1200
         Width           =   2415
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
         TabIndex        =   54
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   2400
         Width           =   1815
      End
   End
   Begin VB.CommandButton CommandClose 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Close"
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
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   120
      Width           =   1695
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
      Height          =   3735
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   3135
      Begin VB.CommandButton CommandMoveTo1 
         Caption         =   "Move To"
         Height          =   495
         Left            =   2280
         TabIndex        =   44
         ToolTipText     =   "Click this button to move to the stage coordinates for point #1 calibration"
         Top             =   3120
         Width           =   735
      End
      Begin VB.CommandButton CommandReadCurrentStageCoordinate1 
         Caption         =   "Read Current Stage Coordinate"
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
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Click this button to load the current stage coordinates for the first calibration point"
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox TextZStage1 
         Height          =   285
         Left            =   2040
         TabIndex        =   41
         ToolTipText     =   "Enter the Z stage coordinate of the first calibration point"
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
         Left            =   240
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Click this button and then click the picture image using the mouse to load the pixel coordinate of the first calibration point"
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox TextYPixel1 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "Click the image to read the Y pixel coordinate of the first calibration point"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox TextXPixel1 
         BackColor       =   &H80000004&
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
         TabIndex        =   53
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
   Begin VB.Label LabelCalibrationAccuracy 
      Alignment       =   2  'Center
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
      Height          =   735
      Left            =   3360
      TabIndex        =   56
      Top             =   4560
      Width           =   2655
   End
   Begin VB.Label LabelCalibration 
      Alignment       =   2  'Center
      Caption         =   "Image Is Calibrated!"
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
      Height          =   495
      Left            =   3720
      TabIndex        =   52
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "When using a three point calibration, the program automatically includes a Z correction for stage sample tilt!"
      Height          =   855
      Left            =   3360
      TabIndex        =   40
      Top             =   8400
      Width           =   2655
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "X and Y pixel coordinates given in ""twip"" units! (15 twips per pixel)"
      Height          =   495
      Left            =   3360
      TabIndex        =   29
      Top             =   6480
      Width           =   2655
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   $"PictureSnap2.frx":008C
      Height          =   1575
      Left            =   3360
      TabIndex        =   28
      Top             =   1440
      Width           =   2655
   End
End
Attribute VB_Name = "FormPICTURESNAP2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2018 by John J. Donovan
Option Explicit

Private Sub CommandCalibratePicture_Click()
If Not DebugMode Then On Error Resume Next
Call PictureSnapCalibrate(Int(0))
If ierror Then Exit Sub
End Sub

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
If Not PictureSnapCalibrated Then
Call PictureSnapCalibrateUnLoad
If ierror Then Exit Sub
End If
Unload FormPICTURESNAP2
Unload FormPICTURESNAP3
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

Private Sub CommandLightModeOff_Click()
If Not DebugMode Then On Error Resume Next
If Light_Reflected_Transmitted& = 0 Then Call RealTimeSetReflectedLightMode(Int(0))
If Light_Reflected_Transmitted& = 1 Then Call RealTimeSetTransmittedLightMode(Int(0))
FormPICTURESNAP2.CommandLightModeOn.BackColor = vbButtonFace
FormPICTURESNAP2.CommandLightModeOff.BackColor = vbWhite
End Sub

Private Sub CommandLightModeOn_Click()
If Not DebugMode Then On Error Resume Next
If Light_Reflected_Transmitted& = 0 Then Call RealTimeSetReflectedLightMode(Int(1))
If Light_Reflected_Transmitted& = 1 Then Call RealTimeSetTransmittedLightMode(Int(1))
FormPICTURESNAP2.CommandLightModeOn.BackColor = vbWhite
FormPICTURESNAP2.CommandLightModeOff.BackColor = vbButtonFace
End Sub

Private Sub CommandLightModeReflected_Click()
If Not DebugMode Then On Error Resume Next
Call RealTimeSetLightMode(Int(0))
FormPICTURESNAP2.CommandLightModeReflected.BackColor = vbWhite
FormPICTURESNAP2.CommandLightModeTransmitted.BackColor = vbButtonFace
End Sub

Private Sub CommandLightModeTransmitted_Click()
If Not DebugMode Then On Error Resume Next
Call RealTimeSetLightMode(Int(1))
FormPICTURESNAP2.CommandLightModeReflected.BackColor = vbButtonFace
FormPICTURESNAP2.CommandLightModeTransmitted.BackColor = vbWhite
End Sub

Private Sub CommandLoadACQCalibration_Click()
If Not DebugMode Then On Error Resume Next
Call PictureSnapLoadACQ
If ierror Then Exit Sub
End Sub

Private Sub CommandMoveTo1_Click()
If Not DebugMode Then On Error Resume Next
Call PictureSnapMoveToCalibrationPoint(Val(FormPICTURESNAP2.TextXStage1), Val(FormPICTURESNAP2.TextYStage1), Val(FormPICTURESNAP2.TextZStage1))
If ierror Then Exit Sub
End Sub

Private Sub CommandMoveTo2_Click()
If Not DebugMode Then On Error Resume Next
Call PictureSnapMoveToCalibrationPoint(Val(FormPICTURESNAP2.TextXStage2), Val(FormPICTURESNAP2.TextYStage2), Val(FormPICTURESNAP2.TextZStage2))
If ierror Then Exit Sub
End Sub

Private Sub CommandMoveTo3_Click()
If Not DebugMode Then On Error Resume Next
Call PictureSnapMoveToCalibrationPoint(Val(FormPICTURESNAP2.TextXStage3), Val(FormPICTURESNAP2.TextYStage3), Val(FormPICTURESNAP2.TextZStage3))
If ierror Then Exit Sub
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

Private Sub CommandReadCurrentStageCoordinate1_Click()
If Not DebugMode Then On Error Resume Next
Call PictureSnapCalibratePointStage(Int(4))
If ierror Then Exit Sub
End Sub

Private Sub CommandReadCurrentStageCoordinate2_Click()
If Not DebugMode Then On Error Resume Next
Call PictureSnapCalibratePointStage(Int(5))
If ierror Then Exit Sub
End Sub

Private Sub CommandReadCurrentStageCoordinate3_Click()
If Not DebugMode Then On Error Resume Next
Call PictureSnapCalibratePointStage(Int(6))
If ierror Then Exit Sub
End Sub

Private Sub Form_Activate()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)   ' save here also because in case user moves window and activate events occurs before closing window

' Just exit if busy, paused or not visible
If RealTimeInterfaceBusy Then Exit Sub
If RealTimePauseAutomation Then Exit Sub
If Not FormPICTURESNAP2.Visible Then Exit Sub

' Just exit if acquisition
If AcquisitionOnMotorCrystal Then Exit Sub
If AcquisitionOnCounterCount Then Exit Sub
If AcquisitionOnSample Then Exit Sub
If AcquisitionOnWavescan Then Exit Sub
If AcquisitionOnPeakCenter Then Exit Sub
If AcquisitionOnAutomate Then Exit Sub
If AcquisitionOnVolatile Then Exit Sub
If AcquisitionOnQuickscan Then Exit Sub
If AcquisitionOnAutoFocus Then Exit Sub

If AcquisitionOnEDS Then Exit Sub
If AcquisitionOnCL Then Exit Sub

' Update the light buttons
If RealTimeMode Then
Call MoveStageMapUpdateButtons(FormPICTURESNAP2)
If ierror Then Exit Sub
End If
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call MiscAlwaysOnTop(True, FormPICTURESNAP2)
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormPICTURESNAP2)
HelpContextID = IOGetHelpContextID("FormPICTURESNAP2")
PictureSnapCalibratedPreviously = PictureSnapCalibrated
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not DebugMode Then On Error Resume Next
If WaitingForCalibrationClick Then
FormPICTURESNAP2.MousePointer = vbDefault
End If
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
