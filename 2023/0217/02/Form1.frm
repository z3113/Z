VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "forѭ����ϰ1"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7965
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
   ScaleHeight     =   4485
   ScaleWidth      =   7965
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "������"
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
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "������"
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "������"
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "������"
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
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "����һ"
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim i%
    Cls
    Form1.Caption = "����һ����ӡ��������"
    Label1.Caption = "1��һ��    �ڴ����ϴ�ӡ1-10֮���������һ��һ����"
    For i = 1 To 10
        Print i
    Next i
End Sub

Private Sub Command2_Click()
    Dim i%, a%
    Cls
    Form1.Caption = "��������������"
    Label1.Caption = "2���ڴ��������50��100֮���������Ҫ��ÿ��5���������Ŀ�ʼ�����"
    For i = 100 To 50 Step -1
        If i Mod 2 <> 0 Then
            Print i;
            a = a + 1
            If a Mod 5 = 0 Then Print
        End If
    Next i
End Sub

Private Sub Command3_Click()
    Dim i%, a%
    Cls
    Form1.Caption = "�����������"
    Label1.Caption = "3����1+2+3+4+��+50֮�͡�"
    For i = 1 To 50
        a = a + i
    Next i
    Print "1+2+3+4+��+50֮��Ϊ��" & a
End Sub

Private Sub Command4_Click()
    Dim i%
    Dim a
    Cls
    Form1.Caption = "�����ģ��׳���ϰ"
    Label1.Caption = "4����20��"
    a = 1
    For i = 1 To 20
        a = a * i
    Next i
    Print "20!="; a
End Sub

Private Sub Command5_Click()
    Dim i%, a%, b%
    Cls
    Form1.Caption = "�����壺���ż������"
    Label1.Caption = "5������Ļ�ϴ�ӡ���1-100֮�������ż����ÿ�����������������֮�͡�"
    For i = 1 To 100
        If i Mod 2 = 0 Then
            Print i;
            a = a + 1
            If a Mod 5 = 0 Then Print
            b = b + i
        End If
    Next i
    Print "ż��֮��Ϊ��"; b
End Sub

Private Sub Command6_Click()
    Form1.Hide
    Form2.Show
End Sub

Private Sub Form_Activate()
    Form1.Caption = "forѭ����ϰ1"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("�Ƿ��˳�", vbYesNo + 64, "�رճ���") = vbNo Then Cancel = True
End Sub
