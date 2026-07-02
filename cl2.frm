VERSION 5.00
Begin VB.Form FormCL 
   Caption         =   "Dummy form for Standard and TestCL"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameXUnits 
      Caption         =   "X Axis Units"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.OptionButton OptionXAxisUnits 
         Caption         =   "Nanometers"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton OptionXAxisUnits 
         Caption         =   "Electron Volts"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "FormCL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
