VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
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
      Height          =   975
      Left            =   6000
      TabIndex        =   7
      Top             =   3000
      Width           =   1695
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
      Height          =   975
      Left            =   1080
      TabIndex        =   6
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton CMDTIME 
      Caption         =   "TIME"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   5
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton CMDDATE 
      Caption         =   "DATE"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6600
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox TXTTIME 
      Height          =   615
      Left            =   3840
      TabIndex        =   3
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox TXTDATE 
      Height          =   735
      Left            =   3840
      TabIndex        =   2
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "TODAYS TIME"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "TODAYS DATE IS"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   2775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDCLEAR_Click()
TXTDATE.Text = ""
TXTTIME.Text = ""
End Sub

Private Sub CMDDATE_Click()
TXTDATE.Text = Date
End Sub

Private Sub CMDEXIT_Click()
End
End Sub

Private Sub CMDTIME_Click()
TXTTIME.Text = Time
End Sub
