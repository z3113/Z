VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "forѭ���ַ�����ϰ"
   ClientHeight    =   7350
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10095
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
   ScaleHeight     =   7350
   ScaleWidth      =   10095
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFC0C0&
      Cancel          =   -1  'True
      Caption         =   "�˳�(ESC)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFC0C0&
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
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFC0C0&
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
      Height          =   615
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0C0&
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
      Height          =   615
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
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
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
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
      Height          =   615
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
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
      Height          =   615
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "����һ(ENTER)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6120
      TabIndex        =   6
      Text            =   "�����ַ�������"
      Top             =   4200
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1215
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   6000
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "Form1.frx":0000
      Top             =   4200
      Width           =   4335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "�����жϣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4920
      TabIndex        =   5
      Top             =   4200
      Width           =   1125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "����/���ܺ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   3
      Top             =   5640
      Width           =   1470
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "����/����ǰ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   1
      Top             =   3840
      Width           =   1470
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   4095
      Left            =   4920
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Cls
    Dim i%, a%, b$, c%
    Label1.Caption = "һ�������ı���������һ���ַ�����ͳ�����е������ж��ٸ���"
    a = Len(Text1.Text)
    For i = 1 To a
        b = Mid(Text1.Text, i, 1)
        If b >= "0" And b <= "9" Then c = c + 1
    Next i
    Print "��������"; c; "��"
End Sub

Private Sub Command2_Click()
    Cls
    Dim i%, a%, b$, c%, d%, e%, f%
    Label1.Caption = "�������ı���������һ���ַ�����ͳ�����еĴ�д��ĸ��Сд��ĸ�����ּ������ַ����ж��ٸ�����Ϣ�������"
    a = Len(Text1.Text)
    For i = 1 To a
        b = Mid(Text1.Text, i, 1)
        If "A" <= b And b <= "Z" Then
            c = c + 1
        ElseIf "a" <= b And b <= "z" Then
            d = d + 1
        ElseIf "0" <= b And b <= "9" Then
            e = e + 1
        Else
            f = f + 1
        End If
    Next i
    Print "��д��ĸ�У�"; c; "��"
    Print "Сд��ĸ�У�"; d; "��"
    Print "�����У�"; e; "��"
    Print "�����ַ��У�"; f; "��"
    MsgBox "��д�У�" & c & vbCrLf & "Сд�У�" & d & vbCrLf & "�����У�" & e & vbCrLf & "�����ַ��У�" & f, vbOKOnly + 64, "ͳ�Ƹ���"
End Sub

Private Sub Command3_Click()
    Cls
    Dim i%, a%, b$, c$
    Label1.Caption = "���������ı����е��ַ����������������ĸΪ��д���ı������޿ո�"
    a = Len(Text1.Text)
    c = UCase(Mid(Text1.Text, 1, 1))
    For i = 2 To a
        b = Mid(Text1.Text, i, 1)
        If b <> " " Then c = c & b
    Next i
    Text2.Text = c
End Sub

Private Sub Command4_Click()
    Cls
    Dim i%, a%, b$, c$
    Label1.Caption = "�ġ������ı����е��ַ���������������ֺ������ַ����䣬��д��ĸ��Сд��Сд��ĸ�Ĵ�д��"
    a = Len(Text1.Text)
    For i = 1 To a
        b = Mid(Text1.Text, i, 1)
        If "A" <= b And b <= "Z" Then
            c = c & LCase(b)
        ElseIf "a" <= b And b <= "z" Then
            c = c & UCase(b)
        Else
            c = c & b
        End If
    Next i
    Text2.Text = c
End Sub

Private Sub Command5_Click()
    Cls
    Dim i%, a$, b%, c%
    Label1.Caption = "�塢����5���ַ������ҳ�A��a��ͷ���ַ���������B��b�������ַ���������"
    For i = 1 To 5
        a = InputBox("�������" & i & "���ַ���", "����ͳ��", "aaaaaa")
        If Mid(a, 1, 1) = "A" Or Mid(a, 1, 1) = "a" Then b = b + 1: Print a
        If Mid(a, Len(a), 1) = "A" Or Mid(a, Len(a), 1) = "b" Then c = c + 1: Print a
    Next i
    Print "A��ͷ���ַ�������"; b
    Print "B�������ַ�������"; c
End Sub

Private Sub Command6_Click()
    Cls
    Dim i%, a%, b$, c$
    Label1.Caption = "��ĸ���ܣ�A/a��C/c B/b��D/d �������� X/x��Z/z Y/y��A/a Z/z��B/b 0��2 1��3...9��1"
    a = Len(Text1.Text)
    For i = 1 To a
        b = Mid(Text1.Text, i, 1)
        If b = " " Or b = Chr(10) Or b = Chr(13) Then
            b = b
        ElseIf ("y" <= b And b <= "z") Or ("Y" <= b And b <= "Z") Then
            b = Chr(Asc(b) - 24)
        ElseIf "8" <= b And b <= "9" Then
            b = Chr(Asc(b) - 8)
        Else
            b = Chr(Asc(b) + 2)
        End If
        c = c & b
    Next i
    Text2.Text = c
End Sub

Private Sub Command7_Click()
    Cls
    Dim i%, a%, b$, c$
    Label1.Caption = "�ߡ����ı���������һ���ַ������ж����Ƿ��ǻ��ģ������ԭ�ַ����͵������ַ�����"
    a = Len(Text3.Text)
    For i = a To 1 Step -1
        b = Mid(Text3.Text, i, 1)
        c = c & b
    Next i
    If c <> Text3.Text Then
        Print "ԭ�ַ���Ϊ��" & Text3.Text
        Print "�������ַ���Ϊ��"; c
        Print "���ǻ���"
    ElseIf c = Text3.Text Then
        Print "ԭ�ַ���Ϊ��" & Text3.Text
        Print "�������ַ���Ϊ��"; c
        Print "�ǻ���"
    End If
End Sub
