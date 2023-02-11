VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Discount"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7575
   LinkTopic       =   "Form4"
   ScaleHeight     =   4065
   ScaleWidth      =   7575
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Unload"
      Height          =   495
      Left            =   6240
      TabIndex        =   0
      Top             =   3240
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a, b

Private Sub Command1_Click()
    Unload Form4
End Sub

Private Sub Form_Activate()
    a = Val(InputBox("请输入购买的单价", "单价"))
    b = Val(InputBox("请输入购买的数量", "数量"))
    If a * b <= 1000 Then
        Print "你购买了单价为" & a & "元的物品" & b & "个，你需付总价为"; a * b & "元，优惠后实际需付" & a * b & "元"
    ElseIf a * b <= 2000 Then
        Print "你购买了单价为" & a & "元的物品" & b & "个，你需付总价为"; a * b & "元，优惠后实际需付" & a * b * 0.9 & "元"
    ElseIf a * b <= 3000 Then
        Print "你购买了单价为" & a & "元的物品" & b & "个，你需付总价为"; a * b & "元，优惠后实际需付" & a * b * 0.8 & "元"
    ElseIf a * b > 3000 Then
        Print "你购买了单价为" & a & "元的物品" & b & "个，你需付总价为"; a * b & "元，优惠后实际需付" & a * b * 0.7 & "元"
    End If
End Sub
