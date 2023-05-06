VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form6"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "CHECK"
      Height          =   975
      Left            =   3000
      TabIndex        =   1
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox TXTINPUT 
      Height          =   975
      Left            =   5280
      TabIndex        =   0
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label LBLRESULT 
      Height          =   615
      Left            =   2880
      TabIndex        =   3
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "ENTER A YEAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   2
      Top             =   480
      Width           =   3615
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
A = Val(TXTINPUT.Text)
If A Mod 4 = 0 Then
    LBLRESULT.Caption = "THIS IS A LEAP YEAR"
Else
    LBLRESULT.Caption = "THIS IS NOT A  LEAP YEAR"
End If
End Sub
