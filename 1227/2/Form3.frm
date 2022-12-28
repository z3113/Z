VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "配置单生成器"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6150
   LinkTopic       =   "Form3"
   ScaleHeight     =   4350
   ScaleWidth      =   6150
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Height          =   2655
      Left            =   3840
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "重选"
      Height          =   495
      Left            =   4800
      TabIndex        =   17
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "确定"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3480
      TabIndex        =   16
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "返回"
      Height          =   495
      Left            =   1800
      TabIndex        =   15
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      Caption         =   "输入CPU型号"
      Height          =   975
      Left            =   1680
      TabIndex        =   13
      Top             =   2400
      Width           =   1935
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "其他"
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   1335
      Begin VB.CheckBox Check3 
         Caption         =   "打印机"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   1200
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         Caption         =   "声卡"
         Height          =   300
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Modem"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "内存"
      Height          =   1935
      Left            =   2040
      TabIndex        =   6
      Top             =   240
      Width           =   1335
      Begin VB.OptionButton Option3 
         Caption         =   "8GB"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   1320
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "4GB"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   840
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "2GB"
         Height          =   375
         Left            =   360
         TabIndex        =   0
         Top             =   360
         Value           =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "品牌"
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "DELL"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "IBM"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "联想"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "方正"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "您选择的配置菜单："
      Height          =   255
      Left            =   3840
      TabIndex        =   19
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a$, b$, c$, d$, e$
Option Explicit

Private Sub Check1_Click()
    If Check1.Value = 1 Then c = vbCrLf & "其他：" & Check1.Caption Else c = ""
End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then d = vbCrLf & "其他：" & Check2.Caption Else d = ""
End Sub

Private Sub Check3_Click()
    If Check3.Value = 1 Then e = vbCrLf & "其他：" & Check3.Caption Else e = ""
End Sub

Private Sub Command1_Click()
    Unload Form2
    Form1.Show
End Sub

Private Sub Command2_Click()
    Text2.Text = a & b & c & d & e & vbCrLf & "CPU型号：" & Text1.Text
End Sub

Private Sub Command3_Click()
    Text1.Text = ""
    Text2.Text = ""
    Label1.BackColor = vbWhite
    Label2.BackColor = vbWhite
    Label3.BackColor = vbWhite
    Label4.BackColor = vbWhite
    Option1.Value = False
    Option2.Value = False
    Option3.Value = False
    Check1.Value = 0
    Check2.Value = 0
    Check3.Value = 0
End Sub

Private Sub Label1_Click()
    Label1.BackColor = vbBlue
    Label2.BackColor = vbWhite
    Label3.BackColor = vbWhite
    Label4.BackColor = vbWhite
    a = "品牌：" & Label1.Caption
End Sub

Private Sub Label2_Click()
    Label1.BackColor = vbWhite
    Label2.BackColor = vbBlue
    Label3.BackColor = vbWhite
    Label4.BackColor = vbWhite
    a = "品牌：" & Label2.Caption
End Sub

Private Sub Label3_Click()
    Label1.BackColor = vbWhite
    Label2.BackColor = vbWhite
    Label3.BackColor = vbBlue
    Label4.BackColor = vbWhite
    a = "品牌：" & Label3.Caption
End Sub

Private Sub Label4_Click()
    Label1.BackColor = vbWhite
    Label2.BackColor = vbWhite
    Label3.BackColor = vbWhite
    Label4.BackColor = vbBlue
    a = "品牌：" & Label4.Caption
End Sub

Private Sub Option1_Click()
    If Option1.Value = True Then b = vbCrLf & "内存：" & Option1.Caption Else b = ""
End Sub

Private Sub Option2_Click()
    If Option2.Value = True Then b = vbCrLf & "内存：" & Option2.Caption Else b = ""
End Sub

Private Sub Option3_Click()
    If Option3.Value = True Then b = vbCrLf & "内存：" & Option3.Caption Else b = ""
End Sub

Private Sub Text1_Change()
    b = vbCrLf & "内存：" & Option1.Caption
    If Text1.Text <> "" Then Command2.Enabled = True
End Sub
