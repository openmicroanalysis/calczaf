VERSION 5.00
Begin VB.Form FormGETZAF 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ZAF, Phi (pz) and Characteristic Fluorescence Selections"
   ClientHeight    =   8520
   ClientLeft      =   1305
   ClientTop       =   975
   ClientWidth     =   18375
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8520
   ScaleWidth      =   18375
   Begin VB.TextBox TextZFractionBackscatterExponent 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4320
      TabIndex        =   81
      ToolTipText     =   $"GETZAF.frx":0000
      Top             =   3450
      Width           =   735
   End
   Begin VB.OptionButton OptionZaf 
      Caption         =   "OptionZaf"
      Height          =   255
      Index           =   11
      Left            =   720
      TabIndex        =   78
      ToolTipText     =   "Donovan and Moy (Armstrong/DAM backscatter) (modified PAP)"
      Top             =   3240
      Width           =   4935
   End
   Begin VB.OptionButton OptionZaf 
      Caption         =   "OptionZaf"
      Height          =   255
      Index           =   10
      Left            =   720
      TabIndex        =   62
      ToolTipText     =   "Pouchou and Pichoir-Simplified Phi-Rho-Z (XPP)"
      Top             =   3000
      Width           =   4935
   End
   Begin VB.OptionButton OptionZaf 
      Caption         =   "OptionZaf"
      Height          =   255
      Index           =   9
      Left            =   720
      TabIndex        =   61
      ToolTipText     =   "Pouchou and Pichoir-Full Phi-Rho-Z (PAP)"
      Top             =   2760
      Width           =   4935
   End
   Begin VB.OptionButton OptionZaf 
      Caption         =   "OptionZaf"
      Height          =   255
      Index           =   8
      Left            =   720
      TabIndex        =   60
      ToolTipText     =   "Bastin PROZA Phi-Rho-Z (EPQ-91)"
      Top             =   2400
      Width           =   4935
   End
   Begin VB.OptionButton OptionZaf 
      Caption         =   "OptionZaf"
      Height          =   255
      Index           =   7
      Left            =   720
      TabIndex        =   59
      ToolTipText     =   "Bastin (original) Phi-Rho-Z"
      Top             =   2160
      Width           =   4935
   End
   Begin VB.OptionButton OptionZaf 
      Caption         =   "OptionZaf"
      Height          =   255
      Index           =   6
      Left            =   720
      TabIndex        =   58
      ToolTipText     =   "Packwood Phi Phi-Rho-Z (EPQ-91)"
      Top             =   1800
      Width           =   4935
   End
   Begin VB.OptionButton OptionZaf 
      Caption         =   "OptionZaf"
      Height          =   255
      Index           =   5
      Left            =   720
      TabIndex        =   57
      ToolTipText     =   "Love-Scott II (Historical ZAF)"
      Top             =   1560
      Width           =   4935
   End
   Begin VB.OptionButton OptionZaf 
      Caption         =   "OptionZaf"
      Height          =   255
      Index           =   4
      Left            =   720
      TabIndex        =   56
      ToolTipText     =   "Love-Scott I (Historical ZAF)"
      Top             =   1320
      Width           =   4935
   End
   Begin VB.OptionButton OptionZaf 
      Caption         =   "OptionZaf"
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   55
      ToolTipText     =   "Heinrich/Duncumb-Reed (Historical ZAF)"
      Top             =   1080
      Width           =   4935
   End
   Begin VB.OptionButton OptionZaf 
      Caption         =   "OptionZaf"
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   54
      ToolTipText     =   "Philibert/Duncumb-Reed (Historical ZAF) (FRAME)"
      Top             =   840
      Width           =   4935
   End
   Begin VB.OptionButton OptionZaf 
      Caption         =   "OptionZaf"
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   53
      ToolTipText     =   "Armstrong/Love-Scott Phi-Rho-Z (default)"
      Top             =   600
      Width           =   4935
   End
   Begin VB.OptionButton OptionZaf 
      Caption         =   "OptionZaf"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   52
      ToolTipText     =   "Select the ZAF or Phi-Rho-Z options individually (expert mode)"
      Top             =   240
      Width           =   4935
   End
   Begin VB.Frame Frame6 
      Caption         =   "BackScatter Coefficients"
      ForeColor       =   &H00FF0000&
      Height          =   1695
      Left            =   12360
      TabIndex        =   33
      Top             =   6720
      Width           =   5895
      Begin VB.OptionButton OptionBsc 
         Caption         =   "OptionBsc"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   79
         Top             =   1320
         Width           =   5175
      End
      Begin VB.OptionButton OptionBsc 
         Caption         =   "OptionBsc"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   65
         Top             =   1080
         Width           =   5175
      End
      Begin VB.OptionButton OptionBsc 
         Caption         =   "OptionBsc"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   64
         Top             =   840
         Width           =   5175
      End
      Begin VB.OptionButton OptionBsc 
         Caption         =   "OptionBsc"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   63
         Top             =   600
         Width           =   5175
      End
      Begin VB.OptionButton OptionBsc 
         Caption         =   "OptionBsc"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Width           =   5175
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Fluorescence Corrections"
      ClipControls    =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   2055
      Left            =   6240
      TabIndex        =   29
      Top             =   6360
      Width           =   5895
      Begin VB.OptionButton OptionFlu 
         Caption         =   "OptionFlu"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   1680
         Width           =   5655
      End
      Begin VB.CheckBox CheckUseFluorescenceByBetaLines 
         Caption         =   "Use Fluorescence By Beta Lines"
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
         Left            =   480
         TabIndex        =   75
         Top             =   1320
         Width           =   5295
      End
      Begin VB.OptionButton OptionFlu 
         Caption         =   "OptionFlu"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1080
         Width           =   5655
      End
      Begin VB.OptionButton OptionFlu 
         Caption         =   "OptionFlu"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   840
         Width           =   5655
      End
      Begin VB.OptionButton OptionFlu 
         Caption         =   "OptionFlu"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   600
         Width           =   5655
      End
      Begin VB.OptionButton OptionFlu 
         Caption         =   "OptionFlu"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   360
         Width           =   5655
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Phi Equations"
      ClipControls    =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   2175
      Left            =   12360
      TabIndex        =   28
      Top             =   720
      Width           =   5895
      Begin VB.OptionButton OptionPhi 
         Caption         =   "OptionPhi"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   1800
         Width           =   5295
      End
      Begin VB.OptionButton OptionPhi 
         Caption         =   "OptionPhi"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1560
         Width           =   5295
      End
      Begin VB.OptionButton OptionPhi 
         Caption         =   "OptionPhi"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1320
         Width           =   5295
      End
      Begin VB.OptionButton OptionPhi 
         Caption         =   "OptionPhi"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1080
         Width           =   5295
      End
      Begin VB.OptionButton OptionPhi 
         Caption         =   "OptionPhi"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   840
         Width           =   5295
      End
      Begin VB.OptionButton OptionPhi 
         Caption         =   "OptionPhi"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   600
         Width           =   5295
      End
      Begin VB.OptionButton OptionPhi 
         Caption         =   "OptionPhi"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   360
         Width           =   5295
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Mean Ionization Potential Corrections"
      ClipControls    =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   2655
      Left            =   6240
      TabIndex        =   27
      Top             =   3360
      Width           =   5895
      Begin VB.OptionButton OptionMip 
         Caption         =   "OptionMip"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   2280
         Width           =   5175
      End
      Begin VB.OptionButton OptionMip 
         Caption         =   "OptionMip"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   2040
         Width           =   5295
      End
      Begin VB.OptionButton OptionMip 
         Caption         =   "OptionMip"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1800
         Width           =   4335
      End
      Begin VB.OptionButton OptionMip 
         Caption         =   "OptionMip"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1560
         Width           =   4335
      End
      Begin VB.OptionButton OptionMip 
         Caption         =   "OptionMip"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1320
         Width           =   4335
      End
      Begin VB.OptionButton OptionMip 
         Caption         =   "OptionMip"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1080
         Width           =   5295
      End
      Begin VB.OptionButton OptionMip 
         Caption         =   "OptionMip"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   840
         Width           =   5295
      End
      Begin VB.OptionButton OptionMip 
         Caption         =   "OptionMip"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   600
         Width           =   5295
      End
      Begin VB.OptionButton OptionMip 
         Caption         =   "OptionMip"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   360
         Width           =   5295
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "BackScatter Corrections"
      ClipControls    =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   3135
      Left            =   12360
      TabIndex        =   26
      Top             =   3240
      Width           =   5895
      Begin VB.OptionButton OptionBks 
         Caption         =   "OptionBks"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   2760
         Width           =   5295
      End
      Begin VB.OptionButton OptionBks 
         Caption         =   "OptionBks"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   2520
         Width           =   5295
      End
      Begin VB.OptionButton OptionBks 
         Caption         =   "OptionBks"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   2280
         Width           =   5295
      End
      Begin VB.OptionButton OptionBks 
         Caption         =   "OptionBks"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   2040
         Width           =   5295
      End
      Begin VB.OptionButton OptionBks 
         Caption         =   "OptionBks"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   1800
         Width           =   5295
      End
      Begin VB.OptionButton OptionBks 
         Caption         =   "OptionBks"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   1560
         Width           =   5295
      End
      Begin VB.OptionButton OptionBks 
         Caption         =   "OptionBks"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   1320
         Width           =   5295
      End
      Begin VB.OptionButton OptionBks 
         Caption         =   "OptionBks"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   1080
         Width           =   5295
      End
      Begin VB.OptionButton OptionBks 
         Caption         =   "OptionBks"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   840
         Width           =   5295
      End
      Begin VB.OptionButton OptionBks 
         Caption         =   "OptionBks"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   600
         Width           =   5295
      End
      Begin VB.OptionButton OptionBks 
         Caption         =   "OptionBks"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   360
         Width           =   5295
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Stopping Power Corrections"
      ClipControls    =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Left            =   6240
      TabIndex        =   25
      Top             =   1080
      Width           =   5895
      Begin VB.OptionButton OptionStp 
         Caption         =   "OptionStp"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   1560
         Width           =   5295
      End
      Begin VB.OptionButton OptionStp 
         Caption         =   "OptionStp"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1320
         Width           =   5295
      End
      Begin VB.OptionButton OptionStp 
         Caption         =   "OptionStp"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1080
         Width           =   5295
      End
      Begin VB.OptionButton OptionStp 
         Caption         =   "OptionStp"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   840
         Width           =   5295
      End
      Begin VB.OptionButton OptionStp 
         Caption         =   "OptionStp"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   600
         Width           =   5295
      End
      Begin VB.OptionButton OptionStp 
         Caption         =   "OptionStp"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   360
         Width           =   5295
      End
   End
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   9360
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton CommandOK 
      BackColor       =   &H00C0FFC0&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Absorption Correction"
      ClipControls    =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   4455
      Left            =   240
      TabIndex        =   0
      Top             =   3960
      Width           =   5895
      Begin VB.OptionButton OptionAbs 
         Caption         =   "OptionAbs"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   4080
         Width           =   5655
      End
      Begin VB.OptionButton OptionAbs 
         Caption         =   "OptionAbs"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   3840
         Width           =   5655
      End
      Begin VB.OptionButton OptionAbs 
         Caption         =   "OptionAbs"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   3480
         Width           =   5655
      End
      Begin VB.OptionButton OptionAbs 
         Caption         =   "OptionAbs"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   3240
         Width           =   5655
      End
      Begin VB.OptionButton OptionAbs 
         Caption         =   "OptionAbs"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   2880
         Width           =   5655
      End
      Begin VB.OptionButton OptionAbs 
         Caption         =   "OptionAbs"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2640
         Width           =   5655
      End
      Begin VB.OptionButton OptionAbs 
         Caption         =   "OptionAbs"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2400
         Width           =   5655
      End
      Begin VB.OptionButton OptionAbs 
         Caption         =   "OptionAbs"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2160
         Width           =   5655
      End
      Begin VB.OptionButton OptionAbs 
         Caption         =   "OptionAbs"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1920
         Width           =   5655
      End
      Begin VB.OptionButton OptionAbs 
         Caption         =   "OptionAbs"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1560
         Width           =   5655
      End
      Begin VB.OptionButton OptionAbs 
         Caption         =   "OptionAbs"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1320
         Width           =   5655
      End
      Begin VB.OptionButton OptionAbs 
         Caption         =   "OptionAbs"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1080
         Width           =   5655
      End
      Begin VB.OptionButton OptionAbs 
         Caption         =   "OptionAbs"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   840
         Width           =   5655
      End
      Begin VB.OptionButton OptionAbs 
         Caption         =   "OptionAbs"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   600
         Width           =   5655
      End
      Begin VB.OptionButton OptionAbs 
         Caption         =   "OptionAbs"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Width           =   5655
      End
   End
   Begin VB.Label LabelZFractionBackscatterExponent 
      Caption         =   "Z Fraction Backscatter Exponent"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1440
      TabIndex        =   80
      ToolTipText     =   "Z fraction exponent (Donovan et al., 2023). Enter zero for exponent based on electron beam energy."
      Top             =   3480
      Width           =   2895
   End
End
Attribute VB_Name = "FormGETZAF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2026 by John J. Donovan
Option Explicit

Private Sub CommandCancel_Click()
If Not DebugMode Then On Error Resume Next
Unload FormGETZAF
End Sub

Private Sub CommandOK_Click()
If Not DebugMode Then On Error Resume Next
Call GetZAFSave
If ierror Then Exit Sub
Unload FormGETZAF
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormGETZAF)
HelpContextID = IOGetHelpContextID("FormGETZAF")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub OptionAbs_Click(Index As Integer)
If Not DebugMode Then On Error Resume Next
' Check if ZAF or prz expression selected
Dim i As Integer
For i% = 0 To UBound(phistring$) - 1
If Index% + 1 < 7 Or (Index% + 1 = 12 Or Index% + 1 = 13) Then
FormGETZAF.OptionPhi(i%).Enabled = False
Else
FormGETZAF.OptionPhi(i%).Enabled = True
End If
Next i%
End Sub

Private Sub OptionZaf_Click(Index As Integer)
If Not DebugMode Then On Error Resume Next
Call GetZAFSetZAF
If ierror Then Exit Sub
Call GetZAFSetEnables
If ierror Then Exit Sub
End Sub

Private Sub TextZFractionBackscatterExponent_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub
