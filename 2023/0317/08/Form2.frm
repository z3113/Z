VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "操作题一"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7365
   LinkTopic       =   "Form2"
   ScaleHeight     =   3975
   ScaleWidth      =   7365
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "返回主界面"
      Height          =   615
      Left            =   4080
      TabIndex        =   5
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "字符串处理"
      Height          =   615
      Left            =   1560
      TabIndex        =   4
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "处理后字符串"
      Top             =   1680
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   1
      Text            =   "请输入字符串"
      Top             =   600
      Width           =   4935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "处理后字符串："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "原始字符串："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1710
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim i%, a%, b$, c$
    a = Len(Text1.Text)
    If a <= 10 Then
        MsgBox "请输入合适长度的字符串（大于10个）", vbOKOnly, "输入提示"
        Text1.SetFocus
        Text1.SelLength = a
        Text2.Text = ""
    Else
        For i = 1 To a Step 2
            b = Mid(Text1.Text, i, 1)
            If ("A" <= b And b <= "Z") Or ("a" <= b And b <= "z") Then
                b = b
            ElseIf "0" <= b And b <= "9" Then
                b = "*"
            Else
                b = "￥"
            End If
            c = c & b
        Next i
        Text2.Text = c
    End If
End Sub

Private Sub Command2_Click()
    Unload Form2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
