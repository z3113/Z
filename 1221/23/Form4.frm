VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "任务三、四则运算"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8415
   LinkTopic       =   "Form4"
   ScaleHeight     =   5760
   ScaleWidth      =   8415
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Caption         =   "返回主窗体"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   10
      Top             =   4800
      Width           =   2895
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Select Case解决"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   9
      Top             =   4800
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ELSEIF解决"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   8
      Top             =   3840
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "普通IF解决"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   7
      Top             =   3840
      Width           =   2895
   End
   Begin VB.TextBox Text1 
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
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "随机产生两个0到100的整数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   2
      Top             =   1320
      Width           =   6255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   4560
      TabIndex        =   6
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "运算结果："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "请输入一个符号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   4800
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Randomize
    Label1.Caption = Int(Rnd * 101)
    Label2.Caption = Int(Rnd * 101)
End Sub

Private Sub Command2_Click()
    If Text1.Text = "+" Then Label5.Caption = Val(Label1.Caption) + Val(Label2.Caption)
    If Text1.Text = "-" Then Label5.Caption = Val(Label1.Caption) - Val(Label2.Caption)
    If Text1.Text = "*" Then Label5.Caption = Val(Label1.Caption) * Val(Label2.Caption)
    If Text1.Text = "/" Then Label5.Caption = Val(Label1.Caption) / Val(Label2.Caption)
End Sub

Private Sub Command3_Click()
    If Text1.Text = "+" Then
        Label5.Caption = Val(Label1.Caption) + Val(Label2.Caption)
    ElseIf Text1.Text = "-" Then
        Label5.Caption = Val(Label1.Caption) - Val(Label2.Caption)
    ElseIf Text1.Text = "*" Then
        Label5.Caption = Val(Label1.Caption) * Val(Label2.Caption)
    ElseIf Text1.Text = "/" Then
        Label5.Caption = Val(Label1.Caption) / Val(Label2.Caption)
    End If
End Sub

Private Sub Command4_Click()
    Select Case Text1.Text
        Case "+"
            Label5.Caption = Val(Label1.Caption) + Val(Label2.Caption)
        Case "-"
            Label5.Caption = Val(Label1.Caption) - Val(Label2.Caption)
        Case Is = "*"
            Label5.Caption = Val(Label1.Caption) * Val(Label2.Caption)
        Case Is = "/"
            Label5.Caption = Val(Label1.Caption) / Val(Label2.Caption)
    End Select
End Sub

Private Sub Command5_Click()
    Unload Form4
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
