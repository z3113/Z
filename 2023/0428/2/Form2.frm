VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "��������"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7920
   LinkTopic       =   "Form2"
   ScaleHeight     =   3960
   ScaleWidth      =   7920
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   10
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   1440
      Width           =   855
   End
   Begin VB.CheckBox Check3 
      Caption         =   "�����ײ�30Ԫ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CheckBox Check2 
      Caption         =   "��׼�ײ�23Ԫ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "��ͯ�ײ�18Ԫ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3480
      TabIndex        =   9
      Top             =   2880
      Width           =   210
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3480
      TabIndex        =   8
      Top             =   2160
      Width           =   210
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3480
      TabIndex        =   7
      Top             =   1440
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�˵�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   480
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    If Check1.Value = 1 Then Text1.Enabled = True Else Text1.Enabled = False
End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then Text1.Enabled = True Else Text1.Enabled = False
End Sub

Private Sub Check3_Click()
    If Check3.Value = 1 Then Text1.Enabled = True Else Text1.Enabled = False
End Sub

Private Sub Command1_Click()
    Dim a%, b%, c%
    If Text1.Enabled = True Then a = Int(Text1.Text)
    If Text2.Enabled = True Then b = Int(Text2.Text)
    If Text3.Enabled = True Then c = Int(Text3.Text)
    MsgBox "һ��" & a * 18 + b * 23 + c * 30 & "Ԫ", vbOKCancel + 64, "���"
End Sub

Private Sub Form_Load()
    Form1.Show
End Sub
