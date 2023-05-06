VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CMDCHECK 
      Caption         =   "CLICK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4320
      TabIndex        =   3
      Top             =   5040
      Width           =   2535
   End
   Begin VB.TextBox TXTNUM 
      Height          =   735
      Left            =   5880
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label LBLNUM 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   2880
      TabIndex        =   2
      Top             =   2400
      Width           =   5895
   End
   Begin VB.Label Label1 
      Caption         =   "ENTER ANY NUMBER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDCHECK_Click()
Dim X As Integer
X = Val(TXTNUM.Text)
If X Mod 2 = 0 Then
    LBLNUM.Caption = " NUMBER IS EVEN "
Else
    LBLNUM.Caption = " NUMBER IS ODD "
End If
End Sub
