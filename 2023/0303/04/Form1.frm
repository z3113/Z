VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "for循环练习"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9690
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   9690
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   5880
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   4440
      Width           =   3615
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFC0FF&
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFC0FF&
      Caption         =   "任务八"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFC0FF&
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
      Height          =   495
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFC0FF&
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
      Height          =   495
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0FF&
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
      Height          =   495
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0FF&
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
      Height          =   495
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0FF&
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
      Height          =   495
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0FF&
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
      Height          =   495
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "任务一"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   2175
      Left            =   5880
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim i%, a%
    Cls
    Label1.Caption = "1、输出100到200之间能被7整除的数输出，5个一行。"
    For i = 100 To 200
        If i Mod 7 = 0 Then
            Print i;
            a = a + 1
            If a Mod 5 = 0 Then Print
        End If
    Next i
    Print
    Print "100到200之间能被7整除的数共有"; a; "个"
End Sub

Private Sub Command2_Click()
    Dim i%
    Cls
    Label1.Caption = "2、找出所有的水仙花数。"
    Print "所有水仙花数是：";
    For i = 100 To 999
        If (i \ 100) ^ 3 + (i \ 10 Mod 10) ^ 3 + (i Mod 10) ^ 3 = i Then Print i;
    Next i
End Sub

Private Sub Command3_Click()
    Dim i%, a%, b%, min%, max%
    Cls
    max = 0
    min = 1000
    Label1.Caption = "3、随机产生100个【1，999】之间的整数，求它们的最大值和最小值，并用消息框输出。"
    For i = 1 To 100
        a = Int(Rnd * 999 + 1)
        Print a;
        If i Mod 5 = 0 Then Print
        If a >= max Then max = a
        If a <= min Then min = a
    Next i
    MsgBox "最大值为：" & max & vbCrLf & "最小值为：" & min
End Sub

Private Sub Command4_Click()
    Dim i&, a&
    Cls
    Label1.Caption = "4、从键盘输入一个整数，判定它是否为素数，在窗体上输出相应信息。"
    a = InputBox("请输入一个整数", "输入", 100)
    For i = 2 To a
        If a Mod i = 0 Then Exit For
    Next i
    If a = i Then
        Print a; "是素数"
    Else
        Print a; "不是素数"
    End If
End Sub

Private Sub Command5_Click()
    Dim i%, n%, a%
    Cls
    Label1.Caption = "5、inputbox输入正整数n，在窗体上依次打印数列12、14、16、18……的前n项，每行5个，紧凑打印"
    n = InputBox("请输入n的值", "输入", 10)
    For i = 12 To 10 + 2 * n
        If i Mod 2 = 0 Then
            Print i;
            a = a + 1
            If a Mod 5 = 0 Then Print
        End If
    Next i
End Sub

Private Sub Command6_Click()
    Dim i%, a$, b%
    a = ""
    Label1.Caption = "6、在文本框中依次显示20、17、14、……、2，每行4个。"
    For i = 20 To 2 Step -1
        If (i + 1) Mod 3 = 0 Then
            a = a & " " & i
            b = b + 1
            If b Mod 4 = 0 Then a = a & vbCrLf
        End If
    Next i
    Text1.Text = a
End Sub

Private Sub Command7_Click()
    Dim i%, n%, a!
    Cls
    Label1.Caption = "7、文本框输入项数n，输出以下数列前n项之和。1/2+3/4+5/6+...+(2n-1)/2n"
    n = InputBox("请输入n的值", "输入", 10)
    For i = 1 To n
        a = a + (2 * i - 1) / 2 / i
    Next i
    Print "n="; n; "和为："; a
End Sub

Private Sub Command8_Click()
    Dim i%, n%, a#
    Cls
    Label1.Caption = "8、inputbox输入项数n，输出以下数列的前n项之积。"
    n = InputBox("请输入n的值", "输入", 10)
    a = 1
    For i = 1 To n
        a = a * 2 ^ (i - 1) / (i + 1)
    Next i
    Print "n="; n; "积为："; a
End Sub

Private Sub Command9_Click()
    Unload Form1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("是否关闭？", vbYesNo + 32, "确定") = vbNo Then Cancel = True
End Sub
