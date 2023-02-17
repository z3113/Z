VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "for循环练习1"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7965
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
   ScaleHeight     =   4485
   ScaleWidth      =   7965
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0E0FF&
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
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0E0FF&
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0E0FF&
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0E0FF&
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
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim i%
    Cls
    Form1.Caption = "任务一：打印连续数字"
    Label1.Caption = "1、一、    在窗体上打印1-10之间的整数，一行一个。"
    For i = 1 To 10
        Print i
    Next i
End Sub

Private Sub Command2_Click()
    Dim i%, a%
    Cls
    Form1.Caption = "任务二：输出奇数"
    Label1.Caption = "2、在窗体上输出50～100之间的奇数，要求每行5个，从最大的开始输出。"
    For i = 100 To 50 Step -1
        If i Mod 2 <> 0 Then
            Print i;
            a = a + 1
            If a Mod 5 = 0 Then Print
        End If
    Next i
End Sub

Private Sub Command3_Click()
    Dim i%, a%
    Cls
    Form1.Caption = "任务三：求和"
    Label1.Caption = "3、求1+2+3+4+…+50之和。"
    For i = 1 To 50
        a = a + i
    Next i
    Print "1+2+3+4+…+50之和为：" & a
End Sub

Private Sub Command4_Click()
    Dim i%
    Dim a
    Cls
    Form1.Caption = "任务四：阶乘练习"
    Label1.Caption = "4、求20！"
    a = 1
    For i = 1 To 20
        a = a * i
    Next i
    Print "20!="; a
End Sub

Private Sub Command5_Click()
    Dim i%, a%, b%
    Cls
    Form1.Caption = "任务五：输出偶数及和"
    Label1.Caption = "5、在屏幕上打印输出1-100之间的所有偶数，每行五个，最后输出他们之和。"
    For i = 1 To 100
        If i Mod 2 = 0 Then
            Print i;
            a = a + 1
            If a Mod 5 = 0 Then Print
            b = b + i
        End If
    Next i
    Print "偶数之和为："; b
End Sub

Private Sub Command6_Click()
    Form1.Hide
    Form2.Show
End Sub

Private Sub Form_Activate()
    Form1.Caption = "for循环练习1"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("是否退出", vbYesNo + 64, "关闭程序") = vbNo Then Cancel = True
End Sub
