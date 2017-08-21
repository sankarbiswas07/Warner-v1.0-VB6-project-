VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form feedback_record 
   Caption         =   "MENTOR'S CORNER :"
   ClientHeight    =   5130
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6840
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   Picture         =   "feedback_record.frx":0000
   ScaleHeight     =   5130
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cbodept 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "feedback_record.frx":71E72
      Left            =   2880
      List            =   "feedback_record.frx":71E82
      TabIndex        =   3
      Top             =   2640
      Width           =   3735
   End
   Begin MSComCtl2.DTPicker dtpfrom 
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   960
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777152
      CalendarTitleBackColor=   16776960
      Format          =   104071169
      CurrentDate     =   42532
   End
   Begin VB.TextBox txtmentor 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   3120
      Width           =   3735
   End
   Begin MSComCtl2.DTPicker dtpto 
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   1680
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   12648447
      CalendarTitleBackColor=   33023
      Format          =   104071169
      CurrentDate     =   42532
   End
   Begin VB.Image Image1 
      Height          =   645
      Left            =   2880
      Picture         =   "feedback_record.frx":71EFA
      Top             =   3720
      Width           =   2355
   End
End
Attribute VB_Name = "feedback_record"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim d As Date



Private Sub Form_Load()
dtpfrom.Value = Date
dtpto.Value = Date
con.Open "mentor"
rs.Open "select * from mentoring", con, adOpenDynamic, adLockOptimistic, adCmdText
End Sub

Private Sub Image1_Click()

If Len(txtmentor.Text) <> 7 Then
    MsgBox "Please Enter Your Correct Mentor ID before proceed.THANK YOU."
    txtmentor.Text = ""
    cbodept.Text = ""
    Exit Sub
    End If
    
Dim ex As New Excel.Application
Dim exwb As Excel.Workbook
Dim exst As Excel.Worksheet
Dim exSelection
Set ex = CreateObject("excel.application")
Set exwb = ex.Workbooks.Add
ex.Visible = True
Set exSelection = ex.Selection
Set exst = exwb.Worksheets(1)
exst.Cells(1, 1) = "Mentor ID"
exst.Cells(1, 2) = "Student Code"
exst.Cells(1, 3) = "Student Name"
exst.Cells(1, 4) = "Library"
exst.Cells(1, 5) = "Canteen"
exst.Cells(1, 6) = "Hostel"
exst.Cells(1, 7) = "University Relation"
exst.Cells(1, 8) = "Internet"
exst.Cells(1, 9) = "Comments on Classroom & LAB"
exst.Cells(1, 10) = "Visited Date"
exst.Range("A1:J1").Interior.Color = RGB(59, 179, 73)
i = 1
While rs.EOF = False
If rs(0).Value = txtmentor.Text Then
For d = dtpfrom To dtpto
If rs(9) = d Then
i = i + 1
For j = 1 To 10
exst.Cells(i, j) = rs(j - 1)
Next j
End If
Next d
'rs.MoveNext
'Else
'rs.MoveNext

rs.MoveNext
ElseIf rs(0).Value <> txtmentor.Text Then ' Or rs(9) <> d Then
rs.MoveNext
'Next d
End If
'Next d
Wend
exst.Columns.EntireColumn.AutoFit
exwb.SaveAs ("D:\MyFirst")
txtmentor.Text = ""
MsgBox "Please Check D drive for Excel Sheet"
Unload Me

'ex.Quit
'rs.Close
con.Close

End Sub

Private Sub txtmentor_KeyPress(KeyAscii As Integer)
Dim ch As String

    ch = Chr$(KeyAscii)
    If Not ( _
        (ch >= "0" And ch <= "9") Or _
        (ch = vbBack) _
    ) Then
        ' Cancel the character.
        KeyAscii = 0
    End If
End Sub
