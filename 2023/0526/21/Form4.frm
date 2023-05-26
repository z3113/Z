VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "素数排序"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8895
   LinkTopic       =   "Form4"
   ScaleHeight     =   7305
   ScaleWidth      =   8895
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "素数从小到大排序"
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "素数每行5个"
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "产生40个【10，90】随机整数"
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   6240
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   8415
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2520
      Width           =   4095
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   4560
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2520
      Width           =   4095
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(40) As Integer, b%, c(40) As Integer

Private Sub Command1_Click()
    Dim i%
    Text1.Text = ""
    For i = 1 To 40
        a(i) = Int(Rnd * 81 + 10)
        Text1.Text = Text1.Text & a(i) & " "
        If i Mod 10 = 0 Then Text1.Text = Text1.Text & vbCrLf
    Next i
End Sub

Private Sub Command2_Click()
    Dim i%, j%, d%
    Text2.Text = ""
    b = 1
    For i = 1 To 40
        For j = 2 To a(i)
            If a(i) Mod j = 0 Then Exit For
        Next j
'        If a(i) = j Then
'            Text2.Text = Text2.Text & a(i) & " "
'            b = b + 1
'            If b Mod 5 = 0 Then Text2.Text = Text2.Text & vbCrLf
'        End If
        If a(i) = j Then
            c(b) = a(i)
            b = b + 1
        End If
    Next i
    For i = 1 To b - 1
        Text2.Text = Text2.Text & c(i) & " "
        If i Mod 5 = 0 Then Text2.Text = Text2.Text & vbCrLf
    Next i
End Sub

Private Sub Command3_Click()
    Dim i%, j%, d%
    Text3.Text = ""
    For i = 2 To b - 1
        For j = 1 To b - i
            If c(j) > c(j + 1) Then d = c(j): c(j) = c(j + 1): c(j + 1) = d
        Next j
    Next i
    For i = 1 To b - 1
        Text3.Text = Text3.Text & c(i) & " "
        If i Mod 5 = 0 Then Text3.Text = Text3.Text & vbCrLf
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
