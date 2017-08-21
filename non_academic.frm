VERSION 5.00
Begin VB.Form non_academic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SISTEMA TUTORIA   v1.0"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "non_academic.frx":0000
   ScaleHeight     =   568.728
   ScaleMode       =   0  'User
   ScaleWidth      =   1638.639
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   19320
      Top             =   1440
   End
   Begin VB.TextBox txtname 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6600
      TabIndex        =   1
      Top             =   3960
      Width           =   3015
   End
   Begin VB.TextBox txtcode 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6600
      TabIndex        =   0
      Top             =   2880
      Width           =   3015
   End
   Begin VB.TextBox txtcomment 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   15840
      TabIndex        =   2
      Top             =   3000
      Width           =   2775
   End
   Begin VB.ComboBox cbointernet 
      BackColor       =   &H00FFFFFF&
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
      ItemData        =   "non_academic.frx":144A62
      Left            =   13080
      List            =   "non_academic.frx":144A72
      TabIndex        =   7
      Top             =   3960
      Width           =   2295
   End
   Begin VB.ComboBox cbohostel 
      BackColor       =   &H00FFFFFF&
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
      ItemData        =   "non_academic.frx":144A84
      Left            =   13800
      List            =   "non_academic.frx":144A94
      TabIndex        =   6
      Top             =   2760
      Width           =   1815
   End
   Begin VB.ComboBox cborelation 
      BackColor       =   &H00FFFFFF&
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
      ItemData        =   "non_academic.frx":144AA6
      Left            =   10320
      List            =   "non_academic.frx":144AB6
      TabIndex        =   5
      Top             =   3960
      Width           =   2175
   End
   Begin VB.ComboBox cbocanteen 
      BackColor       =   &H00FFFFFF&
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
      ItemData        =   "non_academic.frx":144AC8
      Left            =   11880
      List            =   "non_academic.frx":144AD8
      TabIndex        =   4
      Top             =   2760
      Width           =   1575
   End
   Begin VB.ComboBox cbolib 
      BackColor       =   &H00FFFFFF&
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
      ItemData        =   "non_academic.frx":144AEA
      Left            =   10080
      List            =   "non_academic.frx":144AFA
      TabIndex        =   3
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label lbltime 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15600
      TabIndex        =   9
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   18840
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   18840
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   8
      Top             =   1320
      Width           =   1815
   End
End
Attribute VB_Name = "non_academic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str1 As String
Dim str2 As String
Dim strdate As String
Dim stryear As String
Dim i As Integer

Private Sub cmdexit_Click()
Unload Me
FlashScreen.Show
End Sub

Private Sub cmdmentor_Click()
feedback_record.Show
End Sub

Private Sub cmdsave_Click()

End Sub

Private Sub Form_Load()
lbldate.Caption = Format(Date, "dd.mm.yyyy")
lbltime.Caption = Time
If con.State = adStateOpen Then
con.Close
End If
con.Open "mentor"
rs.Open "select * from mentoring", con, adOpenDynamic, adLockOptimistic, adCmdText
End Sub

Private Sub Image1_Click()
If Len(txtcode.Text) <> 17 Then
MsgBox "Please Enter Your Correct Student Code before proceed.THANK YOU."
txtcode.Text = ""
txtname.Text = ""
Exit Sub
End If

rs.AddNew
str1 = Mid(txtcode.Text, 12, 2)
str2 = Mid(txtcode.Text, 15, 3)

For i = 1 To 20
If Val(str1) = 14 Then
If Val(str2) = i Then
rs(0) = 2700110
End If
End If
Next

For i = 21 To 40
If Val(str1) = 14 Then
If Val(str2) = i Then
rs(0) = 2700111
End If
End If
Next
 
For i = 40 To 72
If Val(str1) = 14 Then
If Val(str2) = i Then
rs(0) = 2700112
End If
End If
Next

rs(1) = txtcode.Text
rs(2) = txtname.Text
rs(3) = cbolib.Text
rs(4) = cbocanteen.Text
rs(5) = cbohostel.Text
rs(6) = cborelation.Text
rs(7) = cbointernet.Text
rs(8) = txtcomment.Text
rs(9) = Date

MsgBox "Record Updated"
txtcode.Text = ""
txtcomment.Text = ""
cbointernet.Text = ""
cborelation.Text = ""
txtname.Text = ""
cbolib.Text = ""
cbocanteen.Text = ""
cbohostel.Text = ""

rs.MoveLast
End Sub

Private Sub Image2_Click()
Unload Me
FlashScreen.Show
End Sub

Private Sub Timer1_Timer()
lbldate.Caption = Format(Date, "dd.mm.yyyy")
lbltime.Caption = Time
End Sub
Private Sub txtcode_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
 Dim ch As String

    ch = Chr$(KeyAscii)
    If Not ( _
        (ch >= "a" And ch <= "z") Or _
        (ch >= "A" And ch <= "Z") Or _
        (ch >= "0" And ch <= "9") Or _
        (ch = vbBack) Or _
        (ch = "/") _
    ) Then
        ' Cancel the character.
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtname_Change()
strdate = Mid(rs(9), 4, 2)
stryear = Mid(rs(9), 7, 4)
While txtcode.Text <> rs(1)
rs.MoveNext
If rs.EOF = True Then
rs.Close
rs.Open
Exit Sub
End If
Wend

If txtcode.Text = rs(1) And Month(Date) = Val(strdate) And Year(Date) = Val(stryear) Then
MsgBox "You have already given Feedback for this month on " & rs(9) & ".THANK YOU."
    txtcode.Text = ""
    txtname.Text = ""
    Exit Sub
End If

End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
 Dim ch1 As String

    ch1 = Chr$(KeyAscii)
    If Not ( _
        (ch1 >= "a" And ch1 <= "z") Or _
        (ch1 >= "A" And ch1 <= "Z") Or _
        (ch1 = " ") Or _
        (ch1 = vbBack) _
    ) Then
        ' Cancel the character.
        KeyAscii = 0
    End If
 
End Sub
