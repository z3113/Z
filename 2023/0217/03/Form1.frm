VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "forѭ����ϰ2"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12015
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   12015
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   6
      Top             =   3600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "�˳�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   5
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�����ۼ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   4
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "n��֮��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   3
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "n��֮��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   2
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   1
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "��Ŀ��ʾ����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2895
      Left            =   5880
      TabIndex        =   0
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Randomize
    Dim i%
    Cls
    Label1.Caption = "1�������������ť����ӡ20��3λ���������ÿ��5����"
    For i = 1 To 20
        Print Int(Rnd * 900) + 100;
        If i Mod 5 = 0 Then Print
    Next i
End Sub

Private Sub Command2_Click()
    Dim i%, n%, a#
    Cls
    Label1.Caption = "2��ʹ��inputbox��������n�������������ǰn��֮�ͣ��������뱣����λС��1+1/2+1/4+1/8+1/16+����"
    n = Int(InputBox("����������n��ֵ", "����", 10))
    For i = 1 To n
        a = a + 1 / 2 ^ (i - 1)
    Next i
    a = Round(a, 2)
    Print "������������ǣ�"; n
    Print "1+1/2+1/4+1/8+1/16+����="; a
End Sub

Private Sub Command3_Click()
    Dim i%, n%, a#
    Cls
    Label1.Caption = "3��ʹ���ı�����������n�������������ǰn��֮�����������뱣��λС����1/2��*��2/3��*��3/4��*��4/5�� ����"
    Text1.Visible = True
    n = Val(Text1.Text)
    If n > 0 Then
        a = 1
        For i = 1 To n
            a = a * i / (i + 1)
        Next i
        a = Round(a, 4)
        Print "������������ǣ�"; n
        Print "(1/2)*(2/3)*(3/4)*(4/5)����="; a
    End If
End Sub

Private Sub Command4_Click()
    Dim i%, n%, a#, b%
    Cls
    Label1.Caption = "4����������ʱҪ���������������������N����������[1��20]֮��ģ�����1-2/3+3/5-4/7+����+n/(2n-1),�����ǣ�����ʾ������Ϣ"
    n = Int(InputBox("����������n��ֵ", "����", 10))
    If 1 <= n <= 20 Then
        b = 1
        For i = 1 To n
            a = a + b * i / (2 * i + -1)
            b = -b
        Next i
        a = Round(a, 3)
        Print "������������ǣ�"; n
        Print "1-2/3+3/5-4/7+����+n/(2n-1)="; a
    Else
        MsgBox "n���ڹ涨��Χ�ڣ�", vbOKCancel, "��ʾ"
    End If
End Sub

Private Sub Command5_Click()
    Unload Form1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("���Ҫ�˳���", vbYesNo, "��ʾ") = vbNo Then Cancel = True
End Sub
