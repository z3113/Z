VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C0C0FF&
   Caption         =   "for循环练习4"
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10605
   LinkTopic       =   "Form2"
   ScaleHeight     =   7935
   ScaleWidth      =   10605
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0E0FF&
      Cancel          =   -1  'True
      Caption         =   "退出(ESC)"
      Height          =   495
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "任务四&D"
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "任务三&C"
      Height          =   495
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "任务二&B"
      Height          =   495
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "任务一&A"
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6480
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   6135
      Left            =   0
      ScaleHeight     =   6075
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
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
      Height          =   6135
      Left            =   5520
      TabIndex        =   6
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim i%, a%, b%
    Label1.Caption = "一、随机产生20个【10，90】范围的整数，如果该数能同时被4和6整除则在图片框上输出该数的个位数；否则在图片框上输出该数的十位数。"
    Picture1.Cls
    For i = 1 To 20
        a = Int(Rnd * 81) + 10
        If a Mod 4 = 0 And a Mod 6 = 0 Then
            Picture1.Print a; "能被整除，个位数是："; a Mod 10
        Else
            Picture1.Print a; "不能被整除，十位数是："; a \ 10
        End If
    Next i
End Sub

Private Sub Command2_Click()
    Form2.Hide
    Form3.Show
End Sub

Private Sub Command3_Click()
    Form2.Hide
    Form4.Show
End Sub

Private Sub Command4_Click()
    Dim i%, a&, b&, c&, d&
    a = 1
    b = 1
    d = 2
    Picture1.Cls
    Picture1.Print 1; 1;
    Label1.Caption = "四、兔子繁殖子数列：指的是这样一个数列：1、1、2、3、5、8、13、21、34、……输出数列的前30项（每行3个），最后输出前30项之和。"
    For i = 3 To 30
        c = a + b
        d = d + c
        a = b
        b = c
        Picture1.Print c;
        If i Mod 3 = 0 Then Picture1.Print
    Next i
    Picture1.Print "前30项之和为："; d
End Sub

Private Sub Command5_Click()
    Unload Form2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("确定退出吗？", vbOKCancel + 32, "退出提示") = vbCancel Then Cancel = True
End Sub
