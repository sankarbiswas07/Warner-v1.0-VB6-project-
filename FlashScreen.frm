VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FlashScreen 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WARNER   v1.0"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   20400
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FlashScreen.frx":0000
   ScaleHeight     =   4845
   ScaleWidth      =   20400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   100
      Left            =   6600
      TabIndex        =   7
      Top             =   960
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComCtl2.DTPicker dtpfrom 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd-MMMM-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   3480
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tw Cen MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   102694913
      CurrentDate     =   42520
   End
   Begin MSComCtl2.DTPicker dtpto 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd-MMMM-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   3960
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tw Cen MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   102694913
      CurrentDate     =   42520
   End
   Begin VB.ComboBox cbodept 
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "FlashScreen.frx":1B0DC4
      Left            =   3120
      List            =   "FlashScreen.frx":1B0DD7
      TabIndex        =   2
      Top             =   3000
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txtmentor 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6120
      Top             =   2520
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   2160
      Picture         =   "FlashScreen.frx":1B0E70
      Top             =   4440
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "FROM DATE"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   3480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TO  DATE"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   3960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "DEPARTMENT "
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   3000
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   360
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   360
      Top             =   1920
      Width           =   2775
   End
End
Attribute VB_Name = "FlashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim d As Date
Private Sub Form_Load()
ProgressBar1.Visible = False
dtpfrom.Value = Date
dtpto.Value = Date
If con.State = adStateOpen Then
con.Close
End If
con.Open "mentor"
rs.Open "select * from mentoring", con, adOpenDynamic, adLockOptimistic, adCmdText
End Sub

Private Sub Image2_Click()
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Image3.Visible = True
txtmentor.Visible = True
cbodept.Visible = True
dtpto.Visible = True
dtpfrom.Visible = True
End Sub

Private Sub Image1_Click()
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Image3.Visible = False
txtmentor.Visible = False
cbodept.Visible = False
dtpto.Visible = False
dtpfrom.Visible = False

ProgressBar1.Visible = True
Timer1.Enabled = True
End Sub
Private Sub Image3_Click()
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
exwb.SaveAs ("D:\Record")
FlashScreen.Show
MsgBox "Please Check D drive for Excel Sheet"
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Image3.Visible = False
txtmentor.Visible = False
cbodept.Visible = False
dtpto.Visible = False
dtpfrom.Visible = False
txtmentor.Text = ""
cbodept.Text = ""
dtpto.Value = Date
dtpfrom.Value = Date
'ex.Quit
'rs.Close
'con.Close
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
Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1
If ProgressBar1.Value = 99 Then
ProgressBar1.Value = ProgressBar1.Value + 1
non_academic.Show
ProgressBar1.Visible = False
If ProgressBar1.Value >= ProgressBar1.Max Then
ProgressBar1.Value = 0
Timer1.Enabled = False
End If
End If
End Sub

