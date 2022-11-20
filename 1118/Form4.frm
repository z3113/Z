VERSION 5.00
Begin VB.Form Form4 
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9735
   LinkTopic       =   "Form4"
   ScaleHeight     =   4935
   ScaleWidth      =   9735
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Height          =   495
      Left            =   7920
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Height          =   495
      Left            =   7920
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Height          =   495
      Left            =   7920
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   7920
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   5040
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Text1.Text = InputBox("相连的第一个数据，请输入数据：", "相连", 1) + InputBox("相连的第二个数据，请输入数据：", "相连", 2)
End Sub

Private Sub Command2_Click()
    Text2.Text = Val(InputBox("相加的第二个数据，请输入数据：", "相加", 2)) + Val(InputBox("相加的第一个数据，请输入数据：", "相加", 1))
End Sub

Private Sub Command3_Click()
    MsgBox "我是第一个消息框，文本框text1中的内容是：" & Text1.Text, 1 + 48, "输出数据"
End Sub

Private Sub Command4_Click()
    MsgBox "我是第二个消息框，文本框text2中的内容是：" & Chr(10) & Chr(13) & Text2.Text, 3 + 64, "输出数据"
End Sub

Private Sub Command5_Click()
    Cls
    Print "123456789012345678901234567890"
    Print "1212ab"
    Print "123456789012345678901234567890"
    Print "12 12 ab"
    Print "123456789012345678901234567890"
    Print Tab(3); "12ab"; 12; "ab"
    Print "123456789012345678901234567890"
    Print Tab(6); "12ab", "12 12"
    Print "123456789012345678901234567890"
    Print "12ab12ab", 12; "abbb"
    Print
    Print "123456789012345678901234567890"
    Print 12; 12
    Print
    MsgBox "屏幕打印的是什么鬼东西？？" & Chr(10) & Chr(13) & "请你按照此内容编写PRINT语句", 1 + 32, "想一想"
End Sub

Private Sub Command6_Click()
    Cls
    Print "123456789012345678901234567890用字符串类型输出两个文本框的内容"
    Print "("; Text1.Text; ")"; Text2.Text, Text1.Text; Text2.Text
    Print
    Print "123456789012345678901234567890用数值类型输出两个文本框的内容"
    Print Val(Text1.Text); Val(Text2.Text), Val(Text1.Text); Val(Text2.Text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Form4
    Form1.Show
End Sub
