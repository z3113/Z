VERSION 5.00
Begin VB.Form Form3 
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8655
   LinkTopic       =   "Form3"
   ScaleHeight     =   3630
   ScaleWidth      =   8655
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    Dim a!, b!, c!, d!
    a = InputBox("请您输入窗体的宽度！" & Chr(10) & Chr(13) & "宽度是：", "请您输入窗体的宽度", 1000)
    b = InputBox("请您输入窗体的高度！" & vbCrLf & "高度是：", "请您输入窗体的高度", 500)
    c = InputBox("请您输入窗体的左边距！" & vbCrLf & "左边距是：", "请您输入窗体的左边距", 0)
    d = InputBox("请您输入窗体的右边距！" & Chr(10) & Chr(13) & "右边距是：", "请您输入窗体的右边距", 0)
    Form3.Width = Val(a)
    Form3.Height = Val(b)
    Form3.Left = Val(c)
    Form3.Top = Val(d)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Form3
    Form1.Show
End Sub
