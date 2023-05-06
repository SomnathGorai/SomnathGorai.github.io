VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
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
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "DATAENTRY(extra).frx":0000
      DataField       =   "code"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   6120
      TabIndex        =   21
      Top             =   6480
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "code"
      Text            =   ""
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "DATAENTRY(extra).frx":0014
      Height          =   2535
      Left            =   1080
      OleObjectBlob   =   "DATAENTRY(extra).frx":0028
      TabIndex        =   20
      Top             =   6360
      Width           =   3975
   End
   Begin VB.CommandButton cmdlast 
      Caption         =   ">|"
      Height          =   495
      Left            =   8520
      TabIndex        =   19
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   ">"
      Height          =   495
      Left            =   7800
      TabIndex        =   18
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton cmdprev 
      Caption         =   "<"
      Height          =   495
      Left            =   7200
      TabIndex        =   17
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton cmdfirst 
      Caption         =   "|<"
      Height          =   495
      Left            =   6600
      TabIndex        =   16
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox tot_rec 
      Height          =   975
      Left            =   8520
      TabIndex        =   15
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      DataSource      =   "Data1"
      Height          =   315
      Left            =   6480
      TabIndex        =   13
      Top             =   720
      Width           =   2655
   End
   Begin VB.Data Data1 
      Connect         =   "Access"
      DatabaseName    =   "E:\sabanti\emp.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   7920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "abcd"
      Top             =   4800
      Width           =   2775
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
   Begin VB.Label Label5 
      Caption         =   "total records"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      TabIndex        =   14
      Top             =   1560
      Width           =   1695
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
TXTCODE.Enabled = True
TXTCODE.SetFocus
CMDADD.Enabled = False

FIELD_BLANK
End Sub

Private Sub CMDDEL_Click()
OP_FLAG = "D"
CMDDEL.Enabled = False
TXTCODE.Enabled = True

TXTCODE.SetFocus
End Sub

Private Sub CMDEDIT_Click()
OP_FLAG = "E"
CMDEDIT.Enabled = False
TXTCODE.Enabled = True
TXTCODE.SetFocus
Call FIELD_BLANK

End Sub

Private Sub CMDEXIT_Click()
End
End Sub

Private Sub cmdfirst_Click()
If Data1.Recordset.BOF Then
    Data1.Recordset.MoveFirst
Else
    Data1.Recordset.MoveFirst
End If
End Sub

Private Sub cmdlast_Click()
If Data1.Recordset.EOF Then
    Data1.Recordset.MovePREVIOUS
Else
    Data1.Recordset.Movelast
End If
End Sub

Private Sub cmdnext_Click()
If Data1.Recordset.EOF Then
    Data1.Recordset.MovePREVIOUS
Else
    Data1.Recordset.MoveNEXT
End If
End Sub

Private Sub cmdprev_Click()
If Data1.Recordset.BOF Then
    Data1.Recordset.MoveNEXT
Else
    Data1.Recordset.MovePREVIOUS
End If
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
 conf = MsgBox("CONFIRM OPERATION", vbYesNo)
 If conf = vbYes Then
    Data1.Recordset.Update
    Call ena
    CMDUPD.Enabled = False
    CMDADD.Enabled = True
    CMDADD.SetFocus
    ElseIf conf = vbNo Then
           MsgBox ("OPERATION FAILED")

 End If
'Data1.Recordset.Update
Data1.RecordSource = "SELECT * FROM abcd"
Data1.Refresh


End Sub

Private Sub Combo1_Change()

TXTCODE.Text = Combo1.Text
End Sub

Private Sub DBCombo1_Click(Area As Integer)
'TXTCODE.Text = Data1.Recordset.Fields(0)

End Sub

Private Sub Form_Activate()
Me.WindowState = 2
CMDUPD.Enabled = False
Call ena
Call fld_false
CMDADD.SetFocus
End Sub

Private Sub TXTCODE_KeyPress(KeyAscii As Integer)
Dim v As Boolean
If KeyAscii = 13 Then
    c = Trim(TXTCODE.Text)
    src_code = "select * from abcd where code = " & "'" & c & "'"
    MsgBox src_code
    Data1.RecordSource = src_code
    Data1.Refresh
    If c = "" Or IsNumeric(c) Or IsDate(c) Or Len(Trim(c)) > 5 Then
        v = False
        TXTCODE.Text = ""
        MsgBox "blank/invalid type"
        ElseIf Not Data1.Recordset.EOF Then
                If OP_FLAG = "A" Then
                    MsgBox "duplicate code"
                    TXTCODE.Text = ""
                    v = False
                    Call ena
                    ElseIf OP_FLAG = "E" Then
                            v = True
                            TXTCODE.Enabled = False
                            Call FIELD_BLANK
                            Call field_to_text
                            TXTNAME.Enabled = True
                            TXTNAME.SetFocus
                            ElseIf OP_FLAG = "D" Then
                                    v = True
                                    TXTCODE.Enabled = False
                                    Call FIELD_BLANK
                                    Call field_to_text
                                    conf = MsgBox("delete:sure?", vbYesNo)
                                    If conf = vbYes Then
                                        Data1.Recordset.Delete
                                        Call ena
                                        Call FIELD_BLANK
                                        ElseIf conf = vbNo Then
                                                MsgBox "del operation failed"
                                                Call ena
                                                
                                    End If
                End If
              

        ElseIf Data1.Recordset.EOF Then
            If OP_FLAG = "A" Then
                v = True
                TXTCODE.Enabled = False
                TXTNAME.Enabled = True
                TXTNAME.SetFocus
                ElseIf OP_FLAG = "E" Then
                    MsgBox "code:" & TXTCODE.Text & "is not in the table"
                    TXTCODE.Text = ""
                    v = False
                    ElseIf OP_FLAG = "D" Then
                           MsgBox "code:" & TXTCODE.Text & "is not in the table"
                           TXTCODE.Text = ""
                           v = False
            End If
    End If
End If

                
End Sub

Private Sub TXTDOJ_KeyPress(KeyAscii As Integer)
Dim v As Boolean
If KeyAscii = 13 Then
    d = TXTDOJ.Text
    If Not IsDate(d) Then
        v = False
        MsgBox "blank/invalid type"
        TXTDOJ.Text = ""
        Else
          v = True
          TXTSAL.Enabled = True
          TXTSAL.SetFocus
    End If
End If

End Sub

Private Sub TXTNAME_KeyPress(KeyAscii As Integer)
Dim v As Boolean
If KeyAscii = 13 Then
    X = TXTNAME.Text
    If Trim(X) = "" Or IsNumeric(X) Or IsDate(X) Or Len(Trim(X)) > 30 Then
        v = False
        TXTNAME.Text = ""
        MsgBox "blank/invalid type"
        Else
           v = True
           TXTDOJ.Enabled = True
           TXTDOJ.SetFocus
    End If
End If
End Sub

Private Sub TXTSAL_KeyPress(KeyAscii As Integer)
Dim v As Boolean

If KeyAscii = 13 Then
       s = TXTSAL.Text
       If IsNumeric(s) Then
          v = True
          CMDUPD.Enabled = True
          CMDUPD.SetFocus
          Else
           v = False
           MsgBox "blank/invalid type"
           TXTSAL.Text = ""
       End If
End If
End Sub
Sub FIELD_BLANK()
    'TXTCODE = ""
    TXTNAME = ""
    TXTDOJ = ""
    TXTSAL = ""
End Sub

Sub ena()
s = "select * from abcd"
Data1.RecordSource = s
Data1.Refresh
tr = Data1.Recordset.RecordCount
tot_rec.Text = tr
If tr >= 1 Then
    CMDEDIT.Enabled = True
    CMDDEL.Enabled = True
    Else
         CMDEDIT.Enabled = False
         CMDDEL.Enabled = False
End If
Combo1.Clear
While Not Data1.Recordset.EOF
      Combo1.AddItem Data1.Recordset.Fields(0)
      Data1.Recordset.MoveNEXT
Wend
End Sub
Sub rec_pnt_mv()
s = "select * from abcd"
Data1.RecordSource = s
Data1.Refresh
End Sub
Sub fld_false()
TXTCODE.Enabled = False
TXTNAME.Enabled = False
TXTDOJ.Enabled = False
TXTSAL.Enabled = False
End Sub
Sub field_to_text()
TXTNAME.Text = Data1.Recordset.Fields(1)
TXTDOJ.Text = Data1.Recordset.Fields(2)
TXTSAL.Text = Data1.Recordset.Fields(3)
End Sub
