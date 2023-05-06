VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form4"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5160
      TabIndex        =   9
      Top             =   7680
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      TabIndex        =   8
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8160
      TabIndex        =   7
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5400
      TabIndex        =   6
      Top             =   5640
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3120
      TabIndex        =   5
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   720
      TabIndex        =   4
      Top             =   5640
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   1095
      Left            =   5040
      TabIndex        =   3
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   5040
      TabIndex        =   2
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label LBLRESULT 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   5160
      TabIndex        =   11
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "RESULT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   10
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "INPUT 2ND NO."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      TabIndex        =   1
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "INPUT 1ST NO."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
LBLRESULT.Caption = Val(Text1.Text) + Val(Text2.Text)
End Sub

Private Sub Command2_Click()
LBLRESULT.Caption = Val(Text1.Text) - Val(Text2.Text)

End Sub

Private Sub Command3_Click()
LBLRESULT.Caption = Val(Text1.Text) * Val(Text2.Text)

End Sub

Private Sub Command4_Click()
LBLRESULT.Caption = Val(Text1.Text) / Val(Text2.Text)

End Sub

Private Sub Command5_Click()
Text1.Text = ""
Text2.Text = ""
LBLRESULT.Caption = ""
End Sub

Private Sub Command6_Click()
End
End Sub

Private Sub Form_Load()
Me.WindowState = 2
End Sub
