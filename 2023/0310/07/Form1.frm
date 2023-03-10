VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "for循环字符串练习"
   ClientHeight    =   7350
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10095
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   10095
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFC0C0&
      Cancel          =   -1  'True
      Caption         =   "退出(ESC)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "任务七"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "任务六"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "任务五"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "任务四"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "任务三"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "任务二"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "任务一(ENTER)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6120
      TabIndex        =   6
      Text            =   "回文字符串输入"
      Top             =   4200
      Width           =   3615
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
      ForeColor       =   &H00FF0000&
      Height          =   1215
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   6000
      Width           =   4335
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
      Height          =   1215
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "Form1.frx":0000
      Top             =   4200
      Width           =   4335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "回文判断："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4920
      TabIndex        =   5
      Top             =   4200
      Width           =   1125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "整理/加密后："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   3
      Top             =   5640
      Width           =   1470
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "整理/加密前："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   1
      Top             =   3840
      Width           =   1470
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   4095
      Left            =   4920
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Cls
    Dim i%, a%, b$, c%
    Label1.Caption = "一、在上文本框中输入一个字符串，统计其中的数字有多少个。"
    a = Len(Text1.Text)
    For i = 1 To a
        b = Mid(Text1.Text, i, 1)
        If b >= "0" And b <= "9" Then c = c + 1
    Next i
    Print "共有数字"; c; "个"
End Sub

Private Sub Command2_Click()
    Cls
    Dim i%, a%, b$, c%, d%, e%, f%
    Label1.Caption = "二、在文本框中输入一个字符串，统计其中的大写字母、小写字母、数字及其他字符各有多少个在消息框输出。"
    a = Len(Text1.Text)
    For i = 1 To a
        b = Mid(Text1.Text, i, 1)
        If "A" <= b And b <= "Z" Then
            c = c + 1
        ElseIf "a" <= b And b <= "z" Then
            d = d + 1
        ElseIf "0" <= b And b <= "9" Then
            e = e + 1
        Else
            f = f + 1
        End If
    Next i
    Print "大写字母有："; c; "个"
    Print "小写字母有："; d; "个"
    Print "数字有："; e; "个"
    Print "其他字符有："; f; "个"
    MsgBox "大写有：" & c & vbCrLf & "小写有：" & d & vbCrLf & "数字有：" & e & vbCrLf & "其他字符有：" & f, vbOKOnly + 64, "统计个数"
End Sub

Private Sub Command3_Click()
    Cls
    Dim i%, a%, b$, c$
    Label1.Caption = "三、整理文本框中的字符串，整理规则：首字母为大写，文本框内无空格。"
    a = Len(Text1.Text)
    c = UCase(Mid(Text1.Text, 1, 1))
    For i = 2 To a
        b = Mid(Text1.Text, i, 1)
        If b <> " " Then c = c & b
    Next i
    Text2.Text = c
End Sub

Private Sub Command4_Click()
    Cls
    Dim i%, a%, b$, c$
    Label1.Caption = "四、整理文本框中的字符串，整理规则：数字和其他字符不变，大写字母改小写，小写字母改大写。"
    a = Len(Text1.Text)
    For i = 1 To a
        b = Mid(Text1.Text, i, 1)
        If "A" <= b And b <= "Z" Then
            c = c & LCase(b)
        ElseIf "a" <= b And b <= "z" Then
            c = c & UCase(b)
        Else
            c = c & b
        End If
    Next i
    Text2.Text = c
End Sub

Private Sub Command5_Click()
    Cls
    Dim i%, a$, b%, c%
    Label1.Caption = "五、输入5个字符串，找出A或a开头的字符串个数、B或b结束的字符串个数。"
    For i = 1 To 5
        a = InputBox("请输入第" & i & "个字符串", "个数统计", "aaaaaa")
        If Mid(a, 1, 1) = "A" Or Mid(a, 1, 1) = "a" Then b = b + 1: Print a
        If Mid(a, Len(a), 1) = "A" Or Mid(a, Len(a), 1) = "b" Then c = c + 1: Print a
    Next i
    Print "A开头的字符串个数"; b
    Print "B结束的字符串个数"; c
End Sub

Private Sub Command6_Click()
    Cls
    Dim i%, a%, b$, c$
    Label1.Caption = "字母加密：A/a→C/c B/b→D/d ………… X/x→Z/z Y/y→A/a Z/z→B/b 0→2 1→3...9→1"
    a = Len(Text1.Text)
    For i = 1 To a
        b = Mid(Text1.Text, i, 1)
        If b = " " Or b = Chr(10) Or b = Chr(13) Then
            b = b
        ElseIf ("y" <= b And b <= "z") Or ("Y" <= b And b <= "Z") Then
            b = Chr(Asc(b) - 24)
        ElseIf "8" <= b And b <= "9" Then
            b = Chr(Asc(b) - 8)
        Else
            b = Chr(Asc(b) + 2)
        End If
        c = c & b
    Next i
    Text2.Text = c
End Sub

Private Sub Command7_Click()
    Cls
    Dim i%, a%, b$, c$
    Label1.Caption = "七、在文本框中输入一个字符串，判断它是否是回文，并输出原字符串和倒序后的字符串。"
    a = Len(Text3.Text)
    For i = a To 1 Step -1
        b = Mid(Text3.Text, i, 1)
        c = c & b
    Next i
    If c <> Text3.Text Then
        Print "原字符串为：" & Text3.Text
        Print "倒序后的字符串为："; c
        Print "不是回文"
    ElseIf c = Text3.Text Then
        Print "原字符串为：" & Text3.Text
        Print "倒序后的字符串为："; c
        Print "是回文"
    End If
End Sub
