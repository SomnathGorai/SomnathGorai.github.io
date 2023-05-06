VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "PROCDURE AND FUNCTION"
      Height          =   6735
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   8655
      Begin VB.CommandButton Command2 
         Caption         =   "FUNCTION"
         Height          =   855
         Left            =   4800
         TabIndex        =   6
         Top             =   5640
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "PROCDURE"
         Height          =   855
         Left            =   4800
         TabIndex        =   5
         Top             =   4440
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   975
         Left            =   4680
         TabIndex        =   4
         Top             =   2280
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   855
         Left            =   4680
         TabIndex        =   3
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "VALUE - 1"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1920
         TabIndex        =   2
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "VALUE - 2"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1920
         TabIndex        =   1
         Top             =   2400
         Width           =   2295
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call PROC1(Text1.Text, Text2.Text)
End Sub

Private Sub Command2_Click()
X = FUNC1(Text1.Text, Text2.Text)
MsgBox "RESULT = " & X
End Sub

Private Sub Form_Load()
Me.WindowState = 2
End Sub

Sub PROC1(A, B)
    C = Val(A) + Val(B)
    MsgBox "RESULT = " & C
End Sub
Function FUNC1(A, B)
         C = Val(A) + Val(B)
         FUNC1 = C
End Function
