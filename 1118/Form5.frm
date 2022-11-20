VERSION 5.00
Begin VB.Form Form5 
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7350
   LinkTopic       =   "Form5"
   ScaleHeight     =   4935
   ScaleWidth      =   7350
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Height          =   420
      Left            =   5760
      TabIndex        =   1
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   5760
      TabIndex        =   0
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   3855
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Text2.Text = Int(InputBox("请输入文字大小：", "输入数据", 1))
    Text1.FontSize = Val(Text2.Text)
    Form5.FontSize = Val(Text2.Text)
End Sub

Private Sub Command2_Click()
    Cls
    Print "文本框的内容为："
    Print Text1.Text
End Sub

Private Sub Command3_Click()
    Cls
    Print "文本框的文字大小为： " & Text2.Text & " 号"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Form5
    Form1.Show
End Sub

Private Sub Text1_GotFocus()
    MsgBox "文本框1获得焦点，选中默认的全部文字", vbOKOnly + 48, "提示"
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End Sub
