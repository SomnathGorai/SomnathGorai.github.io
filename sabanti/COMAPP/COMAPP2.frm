VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form5"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CMDCAL 
      Caption         =   "CALCULATE"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7680
      TabIndex        =   12
      Top             =   5400
      Width           =   2415
   End
   Begin VB.CommandButton CMDCLEAR 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      TabIndex        =   11
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton CMDEXIT 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7560
      TabIndex        =   10
      Top             =   7560
      Width           =   2175
   End
   Begin VB.TextBox TXTRE 
      Height          =   855
      Left            =   4080
      TabIndex        =   9
      Top             =   5280
      Width           =   2415
   End
   Begin VB.OptionButton OPT4 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9240
      TabIndex        =   7
      Top             =   3720
      Width           =   1815
   End
   Begin VB.OptionButton OPT3 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6600
      TabIndex        =   6
      Top             =   3720
      Width           =   2295
   End
   Begin VB.OptionButton OPT2 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   5
      Top             =   3840
      Width           =   2415
   End
   Begin VB.OptionButton OPT1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   3840
      Width           =   2175
   End
   Begin VB.TextBox TXT2 
      Height          =   855
      Left            =   4080
      TabIndex        =   3
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox TXT1 
      Height          =   735
      Left            =   4080
      TabIndex        =   2
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "RESULT"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   8
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "INP2"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "INP1"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A, B As Integer
Private Sub CMDCAL_Click()
A = Val(TXT1.Text)
B = Val(TXT2.Text)
If OPT1.Value = True Then
    TXTRE.Text = A + B
ElseIf OPT2.Value = True Then
    TXTRE.Text = A - B
ElseIf OPT3.Value = True Then
    TXTRE.Text = A * B
ElseIf OPT4.Value = True Then
    TXTRE.Text = A / B
End If

End Sub

Private Sub CMDCLEAR_Click()
TXT1.Text = ""
TXT2.Text = ""
TXTRE.Text = ""
OPT1.Value = False
OPT2.Value = False
OPT3.Value = False
OPT4.Value = False

End Sub

Private Sub CMDEXIT_Click()
End
End Sub
