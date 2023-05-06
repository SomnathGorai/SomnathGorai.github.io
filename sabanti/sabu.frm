VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9180
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   ScaleHeight     =   9180
   ScaleWidth      =   10740
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   765
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   1349
      ButtonWidth     =   1429
      ButtonHeight    =   1191
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   7
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "wrdnew"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "wrdopen"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "xlsxnew"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "xlsxopen"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "cut"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "copy"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "paste"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   480
      TabIndex        =   2
      Top             =   2280
      Width           =   4215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   840
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   5520
      TabIndex        =   1
      Top             =   2400
      Width           =   4215
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1560
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   23
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "sabu.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "sabu.frx":067A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "sabu.frx":0CF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "sabu.frx":136E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "sabu.frx":19E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "sabu.frx":2062
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "sabu.frx":26DC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WRD As Object
Dim XBOOK As Object
Dim XEXCEL As Object
Private Sub Form_Load()
Me.WindowState = 2
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Index
       Case 1
            Call WRDNEW
       Case 2
            Call WRDOPEN
        Case 3
            Call XLSXNEW
        Case 4
            Call XLSXOPEN
        Case 5
            Call CUT
        Case 6
            Call COPY
        Case 7
            Call PASTE
End Select

End Sub
Sub WRDNEW()
    Set WRD = CreateObject("WORD.APPLICATION")
    WRD.Visible = True
    WRD.Documents.Add
    
End Sub
Sub WRDOPEN()
    Dim X As Object
    Set WRD = CreateObject("WORD.APPLICATION")
    WRD.Visible = True
    COM1.ShowOpen
    WRD.Documents.Open COM1.FileName
    
End Sub
Sub XLSXNEW()
    Set XEXCEL = CreateObject("EXCEL.APPLICATION")
    XEXCEL.Visible = True
    XEXCEL.WORKBOOKS.Add
    
End Sub
Sub XLSXOPEN()
    Set XEXCEL = CreateObject("EXCEL.APPLICATION")
    XEXCEL.Visible = True
    COM1.ShowOpen
    XEXCEL.WORKBOOKS.Open COM1.FileName
End Sub
Sub COPY()
    Clipboard.SetText Text1.SelText
End Sub
Sub PASTE()
Label1.Caption = Clipboard.GetText
End Sub
Sub CUT()
    Clipboard.SetText Text1.SelText
    Text1.Text = ""
    
End Sub

