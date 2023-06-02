VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00C0C0FF&
   Caption         =   "排序任务二"
   ClientHeight    =   6615
   ClientLeft      =   180
   ClientTop       =   555
   ClientWidth     =   6990
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   6990
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Caption         =   "降序（选择）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   6
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "Form3.frx":0000
      Top             =   4560
      Width           =   6015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "升序（冒泡）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   3
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "平均分"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   2
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "及格人数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "输入成绩"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "排序后的结果为："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3960
      Width           =   2175
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a(10) As Double

Private Sub Command1_Click()
    Dim i%
    Cls
    For i = 1 To 10
        a(i) = Val(InputBox("请输入第" & i & "个学生的成绩", "输入成绩", 60))
    Next i
    Print "所有成绩为：";
    For i = 1 To 10
        Print a(i);
    Next i
End Sub

Private Sub Command2_Click()
    Dim i%, b%, c%
    For i = 1 To 10
        If a(i) >= 60 Then b = b + 1 Else c = c + 1
    Next i
    Print
    Print "不及格人数为：" & c
    Print "及格人数为：" & b
End Sub

Private Sub Command3_Click()
    Dim i%, b#, c%
    For i = 1 To 10
        b = b + a(i)
    Next i
    For i = 1 To 10
        If a(i) > b / 10 Then c = c + 1
    Next i
    Print
    Print "总分：" & b
    Print "平均分：" & Round(b / 10, 2)
    Print "高于平均分的人数：" & c
End Sub

Private Sub Command4_Click()
    Text1.Text = ""
    Dim i%, j%, b#
    For i = 1 To 9
        For j = 1 To 10 - i
            If a(j) > a(j + 1) Then b = a(j): a(j) = a(j + 1): a(j + 1) = b
        Next j
    Next i
    For i = 1 To 10
        Text1.Text = Text1.Text & a(i) & " "
        If i Mod 5 = 0 Then Text1.Text = Text1.Text & vbCrLf
    Next i
End Sub

Private Sub Command5_Click()
    Text1.Text = ""
    Dim i%, j%, b%
    For i = 1 To 9
        b = i
        For j = i + 1 To 10
            If a(j) > a(b) Then b = j
        Next j
        If i <> b Then a(0) = a(i): a(i) = a(b): a(b) = a(0)
    Next i
    For i = 1 To 10
        Text1.Text = Text1.Text & a(i) & " "
        If i Mod 5 = 0 Then Text1.Text = Text1.Text & vbCrLf
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
