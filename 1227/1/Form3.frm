VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "简答应用"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6630
   LinkTopic       =   "Form3"
   ScaleHeight     =   4950
   ScaleWidth      =   6630
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   4800
      ScaleHeight     =   435
      ScaleWidth      =   1395
      TabIndex        =   14
      Top             =   1320
      Width           =   1455
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "成绩汇总"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "批改"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   12
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "出题"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   11
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   2760
      Width           =   6375
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   5400
      TabIndex        =   8
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   3600
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1920
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "运算符选择"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      Begin VB.OptionButton Option4 
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   495
      End
      Begin VB.OptionButton Option3 
         Caption         =   "×"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         Caption         =   "－"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "＋"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   7
      Top             =   600
      Width           =   255
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a%, b%, c%

Private Sub Command1_Click()
    Randomize
    Text1.Text = Int(Rnd * 101)
    Text2.Text = Int(Rnd * 101)
    Text3.Text = ""
    Text4.Text = ""
    If Option1.Value = True Then Label2.Caption = Option1.Caption
    If Option2.Value = True Then Label2.Caption = Option2.Caption
    If Option3.Value = True Then Label2.Caption = Option3.Caption
    If Option4.Value = True Then Label2.Caption = Option4.Caption
    Text3.SetFocus
End Sub

Private Sub Command2_Click()
    a = a + 1
    If Label2.Caption = Option1.Caption Then
        If Val(Text1.Text) + Val(Text2.Text) = Val(Text3.Text) Then
            Label3.ForeColor = vbBlue
            Label3.Caption = "正确"
            b = b + 1
        Else
            Label3.ForeColor = vbRed
            Label3.Caption = "错误"
            c = c + 1
        End If
    ElseIf Label2.Caption = Option2.Caption Then
        If Val(Text1.Text) - Val(Text2.Text) = Val(Text3.Text) Then
            Label3.ForeColor = vbBlue
            Label3.Caption = "正确"
            b = b + 1
        Else
            Label3.ForeColor = vbRed
            Label3.Caption = "错误"
            c = c + 1
        End If
    ElseIf Label2.Caption = Option3.Caption Then
        If Val(Text1.Text) * Val(Text2.Text) = Val(Text3.Text) Then
            Label3.ForeColor = vbBlue
            Label3.Caption = "正确"
            b = b + 1
        Else
            Label3.ForeColor = vbRed
            Label3.Caption = "错误"
            c = c + 1
        End If
    ElseIf Label2.Caption = Option4.Caption Then
        If Val(Text2.Text) = 0 Then
            MsgBox "除数为0，无效题，请重新出题！", 0 + 16, "出错"
        ElseIf Val(Text1.Text) / Val(Text2.Text) = Val(Text3.Text) Then
            Label3.ForeColor = vbBlue
            Label3.Caption = "正确"
            b = b + 1
        ElseIf Val(Text1.Text) / Val(Text2.Text) <> Val(Text3.Text) Then
            Label3.ForeColor = vbRed
            Label3.Caption = "错误"
            c = c + 1
        End If
    End If
    Command3.Enabled = True
    Command2.Enabled = False
End Sub

Private Sub Command3_Click()
    Text4.Text = "你的测试情况如下：" & vbCrLf & "共做了" & a & "道题" & vbCrLf & "做对： " & b & "道，做错： " & c & "道" & vbCrLf & "正确率为：" & b / a * 100 & "%"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

Private Sub Text3_Change()
    Command2.Enabled = True
End Sub
