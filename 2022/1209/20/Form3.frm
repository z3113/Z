VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "2小数整数判断"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "double click"
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_DblClick()
    Dim a!
    a = Val(Text1.Text)
    If a = Int(a) Then
        Print "文本框中的 " & a & " 是整数"
    Else
        Print "文本框中的 " & a & " 是偶数"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
