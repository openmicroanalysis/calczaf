VERSION 5.00
Begin VB.Form FormPICTURESNAP3 
   Caption         =   "Picture Snap Full Window View"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   1560
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Keep the ScaleHeight and ScaleWidth Properties of this form (in design mode) equal for proper scaling"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   4335
      Left            =   480
      Stretch         =   -1  'True
      Top             =   840
      Width           =   4335
   End
End
Attribute VB_Name = "FormPICTURESNAP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2018 by John J. Donovan
Option Explicit

Dim BitMapX As Single
Dim BitMapY As Single

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormPICTURESNAP3)
HelpContextID = IOGetHelpContextID("FormPICTURESNAP3")
icancel = False
End Sub

Private Sub Form_Resize()
If Not DebugMode Then On Error Resume Next
Call PictureSnapResizeFullView
If ierror Then Exit Sub
FormPICTURESNAP3.Image1.Move 0, 0, FormPICTURESNAP3.ScaleWidth, FormPICTURESNAP3.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub Image1_DblClick()
If Not DebugMode Then On Error Resume Next
Call PictureSnapStageMove2(BitMapX!, BitMapY!)
If ierror Then Exit Sub
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not DebugMode Then On Error Resume Next
BitMapX! = X!
BitMapY! = Y!   ' store for double-click
End Sub

