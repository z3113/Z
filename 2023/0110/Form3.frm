VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "�����ֺ�"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8565
   LinkTopic       =   "Form3"
   ScaleHeight     =   6405
   ScaleWidth      =   8565
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command3 
      Caption         =   "�˳�"
      Height          =   615
      Left            =   6480
      TabIndex        =   21
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   615
      Left            =   6480
      TabIndex        =   20
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   615
      Left            =   6480
      TabIndex        =   19
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   4320
      Width           =   5655
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      LargeChange     =   5
      Left            =   1440
      Max             =   60
      Min             =   10
      TabIndex        =   15
      Top             =   3840
      Value           =   10
      Width           =   3855
   End
   Begin VB.Frame Frame3 
      Caption         =   "����"
      Height          =   1695
      Left            =   4200
      TabIndex        =   9
      Top             =   1680
      Width           =   1575
      Begin VB.CheckBox Check3 
         Caption         =   "�»���"
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   1200
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         Caption         =   "��б"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   720
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "�Ӵ�"
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "��ɫ"
      Height          =   1695
      Left            =   2160
      TabIndex        =   5
      Top             =   1680
      Width           =   1575
      Begin VB.OptionButton Option6 
         Caption         =   "��ɫ"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option5 
         Caption         =   "��ɫ"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton Option4 
         Caption         =   "��ɫ"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   1200
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "����"
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   1575
      Begin VB.OptionButton Option3 
         Caption         =   "����"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   1200
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "����"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "����"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Form3.frx":0000
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   17
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "60"
      Height          =   255
      Left            =   5400
      TabIndex        =   16
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      Height          =   255
      Left            =   1080
      TabIndex        =   14
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "�ֺţ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3840
      Width           =   735
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    Text1.FontBold = Not Text1.FontBold
End Sub

Private Sub Check2_Click()
    Text1.FontItalic = Not Text1.FontItalic
End Sub

Private Sub Check3_Click()
    Text1.FontUnderline = Not Text1.FontUnderline
End Sub

Private Sub Command1_Click()
    Dim a$, b$, c$, d$
    a = "�������ݣ�" & Text1.Text
    If Option1.Value = True Then b = "���壺" & Option1.Caption Else b = "���壺"
    If Option2.Value = True Then b = "���壺" & Option2.Caption Else b = "���壺"
    If Option3.Value = True Then b = "���壺" & Option3.Caption Else b = "���壺"
    If Option4.Value = True Then c = "������ɫ��" & Option4.Caption Else c = "������ɫ��"
    If Option5.Value = True Then c = "������ɫ��" & Option5.Caption Else c = "������ɫ��"
    If Option6.Value = True Then c = "������ɫ��" & Option6.Caption Else c = "������ɫ��"
    If Check1.Value = 1 Then d = "����Ϊ��" & Check1.Caption Else d = "����Ϊ��"
    If Check2.Value = 1 Then d = d & " " & Check2.Caption Else d = "����Ϊ��"
    If Check3.Value = 1 Then d = d & " " & Check3.Caption Else d = "����Ϊ��"
    Text2.Text = a & vbCrLf & b & vbCrLf & c & vbCrLf & d
End Sub

Private Sub Command2_Click()
    Text1.Text = ""
    Option1.Value = False
    Option2.Value = False
    Option3.Value = False
    Option4.Value = False
    Option5.Value = False
    Option6.Value = False
    Check1.Value = 0
    Check2.Value = 0
    Check3.Value = 0
    HScroll1.Value = 10
End Sub

Private Sub Command3_Click()
    Unload Form3
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

Private Sub HScroll1_Change()
    Label4.Caption = HScroll1.Value
    Text1.FontSize = HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
    Label4.Caption = HScroll1.Value
    Text1.FontSize = HScroll1.Value
End Sub

Private Sub Option1_Click()
    Text1.FontName = "����"
End Sub

Private Sub Option2_Click()
    Text1.FontName = "����"
End Sub

Private Sub Option3_Click()
    Text1.FontName = "����"
End Sub

Private Sub Option4_Click()
    Text1.ForeColor = vbGreen
End Sub

Private Sub Option5_Click()
    Text1.ForeColor = vbBlue
End Sub

Private Sub Option6_Click()
    Text1.ForeColor = vbRed
End Sub

Private Sub Text1_Change()
    Randomize
    Form3.BackColor = RGB(Int(Rnd * 256), Int(Rnd * 256), Int(Rnd * 256))
End Sub
