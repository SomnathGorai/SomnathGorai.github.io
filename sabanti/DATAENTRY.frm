VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
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
   Begin VB.Data Data1 
      Connect         =   "Access"
      DatabaseName    =   "E:\sabanti\emp.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   7320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "abcd"
      Top             =   7920
      Width           =   2775
   End
   Begin MSDBGrid.DBGrid DBGrid 
      Bindings        =   "DATAENTRY.frx":0000
      Height          =   2295
      Left            =   1320
      OleObjectBlob   =   "DATAENTRY.frx":0014
      TabIndex        =   13
      Top             =   6240
      Width           =   5655
   End
   Begin VB.CommandButton CMDEXIT 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6600
      TabIndex        =   12
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton CMDDEL 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5160
      TabIndex        =   11
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton CMDUPD 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      TabIndex        =   10
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton CMDEDIT 
      Caption         =   "EDIT"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      TabIndex        =   9
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton CMDADD 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      TabIndex        =   8
      Top             =   4800
      Width           =   1095
   End
   Begin VB.TextBox TXTSAL 
      Height          =   855
      Left            =   3240
      TabIndex        =   3
      Top             =   3240
      Width           =   2655
   End
   Begin VB.TextBox TXTDOJ 
      Height          =   855
      Left            =   3240
      TabIndex        =   2
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox TXTNAME 
      Height          =   855
      Left            =   3240
      TabIndex        =   1
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox TXTCODE 
      Height          =   735
      Left            =   3240
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SALARY"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   7
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DT OF JOINING"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   6
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EMP NAME"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   5
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EMP CODE"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   4
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OP_FLAG As String
Private Sub CMDADD_Click()
OP_FLAG = "A"
TXTCODE.SetFocus
FIELD_BLANK
End Sub

Private Sub CMDDEL_Click()
ECODE = InputBox("ENTER code TO DELETE THE RECORD")
X = "SELECT * FROM abcd WHERE code = " & "'" & ECODE & "'"
Data1.RecordSource = X
Data1.Refresh
If Not Data1.Recordset.EOF Then
   TXTCODE.Text = Data1.Recordset.Fields(0)
   TXTNAME.Text = Data1.Recordset.Fields(1)
   TXTDOJ.Text = Data1.Recordset.Fields(2)
   TXTSAL.Text = Data1.Recordset.Fields(3)
   CONF = MsgBox("CONFIRM OPERATION", vbYesNo)
   If CONF = vbYes Then
      Data1.Recordset.Delete
      ElseIf CONF = vbNo Then
           MsgBox ("OPERATION FAILED")

   End If
ElseIf Data1.Recordset.EOF Then
       MsgBox "RECORD NOT FOUND"
End If
End Sub

Private Sub CMDEDIT_Click()
OP_FLAG = "E"
ECODE = InputBox("ENTER code TO MODIFY THE RECORD")
X = "SELECT * FROM abcd WHERE code = " & "'" & ECODE & "'"
Data1.RecordSource = X
Data1.Refresh
If Not Data1.Recordset.EOF Then
   TXTCODE.Text = Data1.Recordset.Fields(0)
   TXTNAME.Text = Data1.Recordset.Fields(1)
   TXTDOJ.Text = Data1.Recordset.Fields(2)
   TXTSAL.Text = Data1.Recordset.Fields(3)
   TXTNAME.SetFocus
   ElseIf Data1.Recordset.EOF Then
          MsgBox "SEARCH FAILED"
End If


End Sub

Private Sub CMDEXIT_Click()
End
End Sub

Private Sub CMDUPD_Click()
If UCase(OP_FLAG) = "A" Then
    Data1.Recordset.AddNew
    ElseIf UCase(OP_FLAG) = "E" Then
           Data1.Recordset.Edit
End If
 Data1.Recordset.Fields(0) = TXTCODE.Text
 Data1.Recordset.Fields(1) = TXTNAME.Text
 Data1.Recordset.Fields(2) = TXTDOJ.Text
 Data1.Recordset.Fields(3) = Val(TXTSAL.Text)
 CONF = MsgBox("CONFIRM OPERATION", vbYesNo)
 If CONF = vbYes Then
    Data1.Recordset.Update
    ElseIf CONF = vbNo Then
           MsgBox ("OPERATION FAILED")

 End If
'Data1.Recordset.Update
Data1.RecordSource = "SELECT * FROM abcd"
Data1.Refresh


End Sub

Private Sub Form_Activate()
Me.WindowState = 2
CMDADD.SetFocus
End Sub

Private Sub TXTCODE_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   TXTNAME.SetFocus
End If
End Sub

Private Sub TXTDOJ_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   TXTSAL.SetFocus
End If

End Sub

Private Sub TXTNAME_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   TXTDOJ.SetFocus
End If


End Sub

Private Sub TXTSAL_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   CMDUPD.SetFocus
End If

End Sub
Sub FIELD_BLANK()
    TXTCODE = ""
    TXTNAME = ""
    TXTDOJ = ""
    TXTSAL = ""
End Sub

