VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "forѭ����ϰ"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9690
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   9690
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   5880
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   4440
      Width           =   3615
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFC0FF&
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFC0FF&
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFC0FF&
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFC0FF&
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0FF&
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0FF&
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0FF&
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0FF&
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0FF&
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   2175
      Left            =   5880
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim i%, a%
    Cls
    Label1.Caption = "1�����100��200֮���ܱ�7�������������5��һ�С�"
    For i = 100 To 200
        If i Mod 7 = 0 Then
            Print i;
            a = a + 1
            If a Mod 5 = 0 Then Print
        End If
    Next i
    Print
    Print "100��200֮���ܱ�7������������"; a; "��"
End Sub

Private Sub Command2_Click()
    Dim i%
    Cls
    Label1.Caption = "2���ҳ����е�ˮ�ɻ�����"
    Print "����ˮ�ɻ����ǣ�";
    For i = 100 To 999
        If (i \ 100) ^ 3 + (i \ 10 Mod 10) ^ 3 + (i Mod 10) ^ 3 = i Then Print i;
    Next i
End Sub

Private Sub Command3_Click()
    Dim i%, a%, b%, min%, max%
    Cls
    max = 0
    min = 1000
    Label1.Caption = "3���������100����1��999��֮��������������ǵ����ֵ����Сֵ��������Ϣ�������"
    For i = 1 To 100
        a = Int(Rnd * 999 + 1)
        Print a;
        If i Mod 5 = 0 Then Print
        If a >= max Then max = a
        If a <= min Then min = a
    Next i
    MsgBox "���ֵΪ��" & max & vbCrLf & "��СֵΪ��" & min
End Sub

Private Sub Command4_Click()
    Dim i&, a&
    Cls
    Label1.Caption = "4���Ӽ�������һ���������ж����Ƿ�Ϊ�������ڴ����������Ӧ��Ϣ��"
    a = InputBox("������һ������", "����", 100)
    For i = 2 To a
        If a Mod i = 0 Then Exit For
    Next i
    If a = i Then
        Print a; "������"
    Else
        Print a; "��������"
    End If
End Sub

Private Sub Command5_Click()
    Dim i%, n%, a%
    Cls
    Label1.Caption = "5��inputbox����������n���ڴ��������δ�ӡ����12��14��16��18������ǰn�ÿ��5�������մ�ӡ"
    n = InputBox("������n��ֵ", "����", 10)
    For i = 12 To 10 + 2 * n
        If i Mod 2 = 0 Then
            Print i;
            a = a + 1
            If a Mod 5 = 0 Then Print
        End If
    Next i
End Sub

Private Sub Command6_Click()
    Dim i%, a$, b%
    a = ""
    Label1.Caption = "6�����ı�����������ʾ20��17��14��������2��ÿ��4����"
    For i = 20 To 2 Step -1
        If (i + 1) Mod 3 = 0 Then
            a = a & " " & i
            b = b + 1
            If b Mod 4 = 0 Then a = a & vbCrLf
        End If
    Next i
    Text1.Text = a
End Sub

Private Sub Command7_Click()
    Dim i%, n%, a!
    Cls
    Label1.Caption = "7���ı�����������n�������������ǰn��֮�͡�1/2+3/4+5/6+...+(2n-1)/2n"
    n = InputBox("������n��ֵ", "����", 10)
    For i = 1 To n
        a = a + (2 * i - 1) / 2 / i
    Next i
    Print "n="; n; "��Ϊ��"; a
End Sub

Private Sub Command8_Click()
    Dim i%, n%, a#
    Cls
    Label1.Caption = "8��inputbox��������n������������е�ǰn��֮����"
    n = InputBox("������n��ֵ", "����", 10)
    a = 1
    For i = 1 To n
        a = a * 2 ^ (i - 1) / (i + 1)
    Next i
    Print "n="; n; "��Ϊ��"; a
End Sub

Private Sub Command9_Click()
    Unload Form1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("�Ƿ�رգ�", vbYesNo + 32, "ȷ��") = vbNo Then Cancel = True
End Sub
