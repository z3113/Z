VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "����"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8415
   LinkTopic       =   "Form2"
   ScaleHeight     =   4470
   ScaleWidth      =   8415
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   8
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   5520
      TabIndex        =   7
      Top             =   1440
      Width           =   2295
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         Caption         =   "��Բ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   11
         Top             =   2040
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   10
         Top             =   1320
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   9
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   2880
      TabIndex        =   3
      Top             =   1440
      Width           =   2295
      Begin VB.CheckBox Check1 
         Caption         =   "�Ӵ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   1815
      End
      Begin VB.CheckBox Check2 
         Caption         =   "��б"
         BeginProperty Font 
            Name            =   "����"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CheckBox Check3 
         Caption         =   "�»���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   2040
         Width           =   1815
      End
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      ItemData        =   "Form2.frx":0000
      Left            =   600
      List            =   "Form2.frx":0010
      TabIndex        =   0
      Text            =   "������ɫ"
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   1
      Text            =   "��ˮ��ɽ���ǽ�ɽ��ɽ��"
      Top             =   360
      Width           =   6495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "��ɫ���ã�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
End
Attribute VB_Name = "Form2"
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

Private Sub Combo1_Click()
    If Combo1.Text = "��" Then
        Text1.ForeColor = vbRed
    ElseIf Combo1.Text = "��" Then
        Text1.ForeColor = vbGreen
    ElseIf Combo1.Text = "��" Then
        Text1.ForeColor = vbBlue
    ElseIf Combo1.Text = "��" Then
        Text1.ForeColor = vbBlack
    End If
End Sub

Private Sub Combo1_Scroll()
    If Combo1.Text = "��" Then
        Text1.ForeColor = vbRed
    ElseIf Combo1.Text = "��" Then
        Text1.ForeColor = vbGreen
    ElseIf Combo1.Text = "��" Then
        Text1.ForeColor = vbBlue
    ElseIf Combo1.Text = "��" Then
        Text1.ForeColor = vbBlack
    End If
End Sub

Private Sub Command1_Click()
    Unload Form2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

Private Sub Option1_Click()
    Text1.FontName = "����"
End Sub

Private Sub Option2_Click()
    Text1.FontName = "����"
End Sub

Private Sub Option3_Click()
    Text1.FontName = "��Բ"
End Sub
