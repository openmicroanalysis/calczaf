VERSION 5.00
Begin VB.Form FormPERIODIC2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select KLM Elements"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton CommandClear 
      BackColor       =   &H0080FFFF&
      Caption         =   "Clear Selections"
      Height          =   375
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   103
      ToolTipText     =   "Clear all the current KLM periodic element selections"
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   99
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   98
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   97
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   96
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   98
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   95
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   94
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   93
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   92
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   91
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   90
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   89
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   88
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   87
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   86
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   85
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   84
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   83
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   82
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   81
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   80
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   79
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   78
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   77
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   76
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   75
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   74
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   73
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   72
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   71
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   70
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   69
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   68
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   67
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   66
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   65
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   64
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   63
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   62
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   61
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   60
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   59
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   58
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   57
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   56
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   55
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   54
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   53
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   52
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   51
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   50
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   49
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   48
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   47
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   46
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   45
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   44
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   43
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   42
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   41
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   40
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   39
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   38
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   37
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   36
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   35
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   34
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   33
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   32
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   31
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   30
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   29
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   28
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   27
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   26
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   25
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   24
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   23
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   22
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   21
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   20
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   19
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   18
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   17
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   16
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   15
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   14
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   13
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   12
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   11
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   10
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   9
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   8
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   7
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   6
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   5
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   4
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   3
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   2
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   1
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton CommandElement 
      Appearance      =   0  'Flat
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
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      ToolTipText     =   "Close the Periodic dialog without displaying the element selections"
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton CommandOK 
      BackColor       =   &H0080FFFF&
      Caption         =   "OK"
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Accept the currently selected KLM elements"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Select a single or multiple elements for KLM display"
      Height          =   615
      Left            =   1080
      TabIndex        =   102
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "FormPERIODIC2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2017 by John J. Donovan
Option Explicit

Private Sub CommandCancel_Click()
If Not DebugMode Then On Error Resume Next
Unload FormPERIODIC2
icancelload = True
End Sub

Private Sub CommandClear_Click()
If Not DebugMode Then On Error Resume Next
Call Periodic2Clear
If ierror Then Exit Sub
End Sub

Private Sub CommandElement_Click(Index As Integer)
If Not DebugMode Then On Error Resume Next
Call Periodic2SelectElement(Index% + 1)
If ierror Then Exit Sub
End Sub

Private Sub CommandOK_Click()
' Save the periodic parameters
If Not DebugMode Then On Error Resume Next
Call Periodic2Save
If ierror Then Exit Sub
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
icancelload = False
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormPERIODIC2)
HelpContextID = IOGetHelpContextID("FormPERIODIC2")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub
