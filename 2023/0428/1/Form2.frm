VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "选择结构"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8310
   LinkTopic       =   "Form2"
   ScaleHeight     =   7950
   ScaleWidth      =   8310
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame6 
      Caption         =   "2、订票"
      Height          =   2895
      Left            =   4560
      TabIndex        =   22
      Top             =   4920
      Width           =   3615
      Begin VB.TextBox Text11 
         Height          =   495
         Left            =   1560
         TabIndex        =   25
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Text10 
         Height          =   495
         Left            =   1560
         TabIndex        =   24
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         Caption         =   "显示"
         Enabled         =   0   'False
         Height          =   615
         Left            =   1080
         TabIndex        =   23
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "输入成绩："
         Height          =   180
         Left            =   360
         TabIndex        =   27
         Top             =   600
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "等级："
         Height          =   180
         Left            =   360
         TabIndex        =   26
         Top             =   1320
         Width           =   540
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "5、编程求数字"
      Height          =   2055
      Left            =   4560
      TabIndex        =   19
      Top             =   2520
      Width           =   3615
      Begin VB.TextBox Text9 
         Enabled         =   0   'False
         Height          =   615
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   3135
      End
      Begin VB.CommandButton Command5 
         Caption         =   "计算"
         Height          =   615
         Left            =   1080
         TabIndex        =   20
         Top             =   1200
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "4、所得税"
      Height          =   2055
      Left            =   4560
      TabIndex        =   15
      Top             =   120
      Width           =   3615
      Begin VB.TextBox Text8 
         Height          =   495
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Text7 
         Enabled         =   0   'False
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "计算"
         Height          =   495
         Left            =   2160
         TabIndex        =   16
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "3、求方程根"
      Height          =   1935
      Left            =   120
      TabIndex        =   12
      Top             =   5880
      Width           =   3615
      Begin VB.CommandButton Command3 
         Caption         =   "求根"
         Height          =   495
         Left            =   1320
         TabIndex        =   14
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "2、订票"
      Height          =   3255
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   3615
      Begin VB.CommandButton ddddd 
         Caption         =   "显示"
         Height          =   495
         Left            =   1200
         TabIndex        =   11
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   1320
         TabIndex        =   10
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   1320
         TabIndex        =   8
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   1320
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "优惠率"
         Height          =   180
         Left            =   360
         TabIndex        =   9
         Top             =   1920
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "订票数"
         Height          =   180
         Left            =   360
         TabIndex        =   7
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "月份"
         Height          =   180
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "1、工资调整"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.CommandButton Command1 
         Caption         =   "计算"
         Height          =   495
         Left            =   2160
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim a#
    a = Val(Text1.Text)
    If a >= 1800 Then
        a = a * 1.15
    ElseIf a >= 1300 Then
        a = a * 1.15
    Else
        a = a * 1.1
    End If
    Text2.Text = a
End Sub

Private Sub Command2_Click()
    Dim a%, b%
    a = Int(Val(Text3.Text))
    b = Int(Val(Text4.Text))
    If 7 <= a And a <= 9 Then
        If b > 20 Then Text5.Text = "15%" Else Text5.Text = "5%"
    ElseIf (1 <= a And a <= 5) Or a = 10 Or a = 11 Then
        If b > 20 Then Text5.Text = "30%" Else Text5.Text = "20%"
    Else
        Text5.Text = ""
    End If
End Sub

Private Sub Command3_Click()
    Dim a%, b%, c%, d%, x1#, x2#
    a = Val(InputBox("请输入一元二次方程ax^2+bx+c=0中的a"))
    b = Val(InputBox("请输入一元二次方程ax^2+bx+c=0中的b"))
    c = Val(InputBox("请输入一元二次方程ax^2+bx+c=0中的c"))
    If a <> 0 Then
        d = b ^ 2 - 4 * a * c
        If d > 0 Then
            x1 = (-b + Sqr(d)) / 2 / a
            x2 = (-b - Sqr(d)) / 2 / a
            Text6.Text = "方程有实数根为：" & vbCrLf & x1 & vbCrLf & x2
        ElseIf d = 0 Then
            x1 = -b / 2 / a
            Text6.Text = "方程有实数根为：" & vbCrLf & x1
        Else
            Text6.Text = "方程无实数根"
        End If
    Else
        If b <> 0 Then
            x1 = -c / b
        Else
            Text6.Text = "方程无实数根"
        End If
    End If
    
End Sub

Private Sub Command4_Click()
    Dim a%
    a = Val(Text8.Text)
    If a <= 1000 Then
        Text7.Text = "应纳税款：" & a * 0.03
    ElseIf a <= 3000 Then
        Text7.Text = "应纳税款：" & (a - 800) * 0.15
    ElseIf a <= 5000 Then
        Text7.Text = "应纳税款：" & 330 + (a - 3000) * 0.2
    Else
        Text7.Text = "应纳税款：" & a * 0.25
    End If
End Sub

Private Sub Command5_Click()
    Dim a#, b#
    a = Val(InputBox("请输入a的值"))
    b = Val(InputBox("请输入b的值"))
    If a ^ 2 + b ^ 2 >= 100 Then
        Text9.Text = "a^2 + b^2百位以上的数字为 " & (a ^ 2 + b ^ 2) \ 100
    Else
        Text9.Text = "两数之和 " & a + b
    End If
End Sub

Private Sub Command6_Click()
    Dim a%
    a = Val(Text11.Text)
    If a > 90 Then
        Text10.Text = "A"
    ElseIf a > 80 Then
        Text10.Text = "B"
    ElseIf a > 70 Then
        Text10.Text = "C"
    ElseIf a > 60 Then
        Text10.Text = "D"
    Else
        Text10.Text = "E"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

Private Sub Text11_Change()
    If Text11.Text <> "" Then Command6.Enabled = True Else Command6.Enabled = False
End Sub
