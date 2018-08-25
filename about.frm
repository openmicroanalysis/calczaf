VERSION 5.00
Begin VB.Form FormABOUT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About This Program"
   ClientHeight    =   11865
   ClientLeft      =   2400
   ClientTop       =   1545
   ClientWidth     =   9495
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
   ScaleHeight     =   11865
   ScaleWidth      =   9495
   Begin VB.CommandButton CommandOK 
      BackColor       =   &H00C0FFC0&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   11280
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      ClipControls    =   0   'False
      Height          =   11055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      Begin VB.TextBox TextDisclaimer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   9960
         Width           =   6375
      End
      Begin VB.Label LabelAboutSpecialists 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   2175
         Left            =   120
         TabIndex        =   6
         Top             =   7680
         Width           =   9015
      End
      Begin VB.Label LabelAboutBeta 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   1095
         Left            =   600
         TabIndex        =   5
         Top             =   1920
         Width           =   8055
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   240
         Picture         =   "ABOUT.frx":0000
         Top             =   240
         Width           =   960
      End
      Begin VB.Label LabelAbout 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4455
         Left            =   1320
         TabIndex        =   1
         Top             =   3120
         Width           =   6615
      End
      Begin VB.Label LabelAboutTitle 
         Alignment       =   2  'Center
         Height          =   1455
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   9015
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "www.probesoftware.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1680
      MouseIcon       =   "ABOUT.frx":0701
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "Click here to visit the Probe Software web site"
      Top             =   11400
      Width           =   2775
   End
End
Attribute VB_Name = "FormABOUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2018 by John J. Donovan
Option Explicit

Private Const clrLinkActive& = vbBlue
Private Const clrLinkHot& = vbRed
'Private Const clrLinkInactive& = vbBlack

Private Sub CommandOK_Click()
If Not DebugMode Then On Error Resume Next
Unload FormABOUT
End Sub

Private Sub Form_Activate()
' Load FormABOUT text
If Not DebugMode Then On Error Resume Next
Dim tmsg As String

If UCase$(app.EXEName) = UCase$("Probewin") Then
tmsg$ = "Probe for EPMA v. " & ProgramVersionString$ & vbCrLf
Else
tmsg$ = "Program " & app.EXEName & " v. " & ProgramVersionString$ & vbCrLf
End If

tmsg$ = tmsg$ & "For Windows XP/Vista/Win7/Win8/Win10" & vbCrLf
tmsg$ = tmsg$ & vbCrLf
tmsg$ = tmsg$ & "Written by John J. Donovan, Probe Software, Inc." & vbCrLf
tmsg$ = tmsg$ & "(c) Copyright 1995-2018, All Rights Reserved" & vbCrLf & vbCrLf
tmsg$ = tmsg$ & "Special thanks to Paul Carpenter for his tireless testing and many helpful discussions"
FormABOUT.LabelAboutTitle.Caption = tmsg$

tmsg$ = "Many thanks to our excellent and hard working beta testing team:" & vbCrLf & vbCrLf
tmsg$ = tmsg$ & "Anette von der Handt and Karsten Goemann (JEOL 8530), Owen Neill (JEOL 8500), Paul Carpenter (JEOL 8200), "
tmsg$ = tmsg$ & "Heather Lowers (JEOL 8900), Angus Netting (Cameca SXFIVE), Gareth Seward and Karsten Goemann (Cameca SX100)"
FormABOUT.LabelAboutBeta.Caption = tmsg$

tmsg$ = vbCrLf & "Thanks to Paul Carpenter, Don Snyder and Mark Rivers for their helpful advice and" & vbCrLf
tmsg$ = tmsg$ & "Thanks also to Tracy Tingle and Dan Kremser for beta testing and suggestions" & vbCrLf
tmsg$ = tmsg$ & "Special thanks to John Armstrong for the CITZAF quantitative routines and" & vbCrLf
tmsg$ = tmsg$ & "John Friday and Brian Gaynor for help with the hardware interfacing and" & vbCrLf
tmsg$ = tmsg$ & "Paul Carpenter for help with the JEOL 8900/8200 interfacing and" & vbCrLf
tmsg$ = tmsg$ & "John Fournelle for numerous suggestions and Jennifer Donovan for data entry" & vbCrLf & vbCrLf
tmsg$ = tmsg$ & "Nicholas Ritchie for help with various fundamental parameterizations" & vbCrLf
tmsg$ = tmsg$ & "Paul Wallace for help with the Time Dependent Intensity (TDI) correction" & vbCrLf
tmsg$ = tmsg$ & "Alan Rempel and Dave Schmidt for mathematical advice and support and" & vbCrLf
tmsg$ = tmsg$ & "Francesc Salvat and Xavier Llovet for help with secondary fluorescence" & vbCrLf
tmsg$ = tmsg$ & "Scotty Cornelius and Dave Adams for help with JEOL 8500 and 8900 interfacing" & vbCrLf
tmsg$ = tmsg$ & "Julien Allaz, Mike Williams and Mike Jercinovic for help with off-peak modeling" & vbCrLf
tmsg$ = tmsg$ & "Julie Barkman for Surfer/Grapher scripting and Gareth Seward for programming help" & vbCrLf
tmsg$ = tmsg$ & "Thanks to our consulting statistician Kardi Takenomo for help with image processing and" & vbCrLf
tmsg$ = tmsg$ & "Zack Gainsforth for physics and math consulting and" & vbCrLf
tmsg$ = tmsg$ & "Anette von der Handt for design and feature suggestions and" & vbCrLf
tmsg$ = tmsg$ & "Brian Gaynor for help with all sorts of application and driver development work!" & vbCrLf
tmsg$ = tmsg$ & vbCrLf
tmsg$ = tmsg$ & "For technical support and sales, please contact John Donovan at Probe Software, Inc." & vbCrLf
tmsg$ = tmsg$ & "TEL: (541) 343-3400  or  URL: www.probesoftware.com, donovan@probesoftware.com"
FormABOUT.LabelAbout.Caption = tmsg$

tmsg$ = vbCrLf & "For additional support, consultation and/or training please contact our team of Microprobe Specialists:" & vbCrLf
tmsg$ = tmsg$ & vbCrLf
tmsg$ = tmsg$ & "Paul Carpenter, 314 602-9697, carpenter@probesoftware.com" & vbCrLf
tmsg$ = tmsg$ & "Gareth Seward, 805 637-7265, seward@probesoftware.com" & vbCrLf
tmsg$ = tmsg$ & "Karsten Goemann, +61 407-101-990, goemann@probesoftware.com" & vbCrLf
tmsg$ = tmsg$ & "Owen Neill, 207 653-6331, neill@probesoftware.com" & vbCrLf
tmsg$ = tmsg$ & "Anette von der Handt, 612 222-6711, vonderhandt@probesoftware.com"

FormABOUT.LabelAboutSpecialists.Caption = tmsg$

tmsg$ = "IN NO EVENT SHALL PROBE SOFTWARE BE LIABLE TO ANY PARTY "
tmsg$ = tmsg$ & "FOR DIRECT, INDIRECT, SPECIAL, INICIDENTAL, OR CONSEQUENTIAL DAMAGES, "
tmsg$ = tmsg$ & "INCLUDING LOST PROFITS, ARISING OUT OF THE USE OF THIS SOFTWARE AND ITS "
tmsg$ = tmsg$ & "DOCUMENTATION, EVEN IF PROBE SOFTWARE HAS BEEN ADVISED OF "
tmsg$ = tmsg$ & "THE POSSIBILITY OF SUCH DAMAGE." & vbCrLf & vbCrLf

tmsg$ = tmsg$ & "PROBE SOFTWARE SPECIFICALLY DISCLAIMS ANY WARRANTIES, "
tmsg$ = tmsg$ & "INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF "
tmsg$ = tmsg$ & "MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE. THE SOFTWARE "
tmsg$ = tmsg$ & "PROVIDED  HEREUNDER IS ON AN AS IS BASIS,  AND PROBE SOFTWARE "
tmsg$ = tmsg$ & "HAVE NO OBLIGATIONS TO PROVIDE MAINTENANCE, SUPPORT, UPDATES, "
tmsg$ = tmsg$ & "ENHANCEMENTS, OR MODIFICATIONS."
FormABOUT.TextDisclaimer.Text = tmsg$
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call MiscCenterForm(FormABOUT)
Call MiscLoadIcon(FormABOUT)
HelpContextID = IOGetHelpContextID("FormABOUT")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
End Sub

Private Sub Label1_Click()
If Not DebugMode Then On Error Resume Next
Call IOBrowseHTTP(ProbeSoftwareInternetBrowseMethod%, "https://probesoftware.com/index.html")
If ierror Then Exit Sub
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Not DebugMode Then On Error Resume Next
' When the label is clicked, change the color to indicate it is hot
If FormABOUT.Label1.ForeColor = clrLinkActive& Then
FormABOUT.Label1.ForeColor = clrLinkHot&
FormABOUT.Label1.Refresh
End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Not DebugMode Then On Error Resume Next
' Mouse released, so restore the label to clrLinkActive&
If FormABOUT.Label1.ForeColor = clrLinkHot& Then
FormABOUT.Label1.ForeColor = clrLinkActive&
FormABOUT.Label1.Refresh
End If
End Sub

