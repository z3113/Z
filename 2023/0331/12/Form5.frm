VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "百钱买百鸡"
   ClientHeight    =   3570
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7605
   LinkTopic       =   "Form5"
   ScaleHeight     =   3570
   ScaleWidth      =   7605
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "答案"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "1、公鸡每只5元，母鸡每只3元，小鸡三只一元，问100元买100只鸡，每种鸡至少买一只，可以各买几只？"
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   1680
      Width           =   6375
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Cls
    Print "有下列三种方案："
    Print "公鸡", "母鸡", "小鸡"
    Dim i%, j%
    For i = 1 To 18
        For j = 1 To 32
            If i * 5 + j * 3 + (100 - i - j) / 3 = 100 Then Print i, j, (100 - i - j)
        Next j
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form4.Show
End Sub
