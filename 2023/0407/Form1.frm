VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "forѭ���ۼ��۳���ϰ"
   ClientHeight    =   7110
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12870
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   12870
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame2 
      Caption         =   "˫��ѭ��"
      Height          =   6255
      Left            =   6600
      TabIndex        =   13
      Top             =   480
      Width           =   6135
      Begin VB.CommandButton Command18 
         Caption         =   "��ʮ��"
         Height          =   495
         Left            =   4680
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command17 
         Caption         =   "��ʮһ��"
         Height          =   495
         Left            =   4680
         TabIndex        =   20
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command16 
         Caption         =   "��ʮ����"
         Height          =   495
         Left            =   4680
         TabIndex        =   19
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton Command15 
         Caption         =   "��ʮ����"
         Height          =   495
         Left            =   4680
         TabIndex        =   18
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Command14 
         Caption         =   "��ʮ����"
         Height          =   495
         Left            =   4680
         TabIndex        =   17
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton Command13 
         Caption         =   "��ʮ����"
         Height          =   495
         Left            =   4680
         TabIndex        =   16
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton Command12 
         Caption         =   "��ʮ����"
         Height          =   495
         Left            =   4680
         TabIndex        =   15
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CommandButton Command11 
         Caption         =   "��ʮ����"
         Height          =   495
         Left            =   4680
         TabIndex        =   14
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label Label6 
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
         Height          =   1425
         Left            =   240
         TabIndex        =   24
         Top             =   480
         Width           =   3500
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "������ڣ�"
         Height          =   180
         Left            =   240
         TabIndex        =   23
         Top             =   4800
         Width           =   900
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   500
         Left            =   1200
         TabIndex        =   22
         Top             =   4680
         Width           =   3240
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "����ѭ��"
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6135
      Begin VB.CommandButton Command9 
         Caption         =   "�ھ���"
         Height          =   495
         Left            =   4680
         TabIndex        =   9
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton Command8 
         Caption         =   "�ڰ���"
         Height          =   495
         Left            =   4680
         TabIndex        =   8
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton Command7 
         Caption         =   "������"
         Height          =   495
         Left            =   4680
         TabIndex        =   7
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "������"
         Height          =   495
         Left            =   4680
         TabIndex        =   6
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "������"
         Height          =   495
         Left            =   4680
         TabIndex        =   5
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "������"
         Height          =   495
         Left            =   4680
         TabIndex        =   4
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "������"
         Height          =   495
         Left            =   4680
         TabIndex        =   3
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "�ڶ���"
         Height          =   495
         Left            =   4680
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "��һ��"
         Height          =   495
         Left            =   4680
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   500
         Left            =   1200
         TabIndex        =   12
         Top             =   4680
         Width           =   3240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "������ڣ�"
         Height          =   180
         Left            =   240
         TabIndex        =   11
         Top             =   4800
         Width           =   900
      End
      Begin VB.Label Label1 
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
         Height          =   1425
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   3500
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Label1.Caption = "1����S=1+2+3+����+99"
    Dim i%, a%
    For i = 1 To 99
        a = a + i
    Next i
    Label3.Caption = a
End Sub

Private Sub Command11_Click()
    Label6.Caption = "17����������ʱҪ���������������������N���������ǡ�1,20��֮��ģ�����1-2/3!+3/5!-4/7!+��+n/(2*n-1)!�������ǣ�����ִ�����ʾ��Ϣ��"
    Dim i%, j%, a#, b#, n%
    n = Val(InputBox("������n��ֵ", "����", 10))
    If 1 <= n And n <= 20 Then
        For i = 1 To n
            a = i
            For j = 1 To 2 * i - 1
                a = a / j
            Next j
            If i Mod 2 = 0 Then b = b - a Else b = b + a
        Next i
        Label4.Caption = b
    Else
        MsgBox "��������ݲ�����Ҫ��", vbOKCancel + 16, "������ʾ"
    End If
End Sub

Private Sub Command12_Click()
    Label6.Caption = "16����s=1+(2-4)+(2-4+6)+(2-4+6-8)+����(2-4+6-8��)ǰn���ֵ"
    Dim i%, j%, a#, b#, n%
    n = Val(InputBox("������n��ֵ", "����", 10))
    a = 1
    For i = 2 To n
        b = 0
        For j = 2 To 2 * i Step 2
            If j Mod 4 = 0 Then b = b - j Else b = b + j
        Next j
        a = a + b
    Next i
    Label4.Caption = a
End Sub

Private Sub Command13_Click()
    Label6.Caption = "15����s=1+(1*3)-(1*3*5)+(1*3*5*7)+��(1*3*5*7*��*(2*n-1))"
    Dim i%, j%, a#, b#, n%
    n = Val(InputBox("������n��ֵ", "����", 10))
    b = 1
    For i = 2 To n
        a = 1
        For j = 1 To 2 * i - 1 Step 2
            a = a * j
        Next j
        If i Mod 2 = 0 Then b = b + a Else b = b - a
    Next i
    Label4.Caption = b
End Sub

Private Sub Command14_Click()
    Label6.Caption = "14����S=1-1/2!+1/3!+��+1/2n!"
    Dim i%, j%, a#, b#, n%
    n = Val(InputBox("������n��ֵ", "����", 10))
    For i = 1 To 2 * n
        a = 1
        For j = 1 To i
            a = a * j
        Next j
        If i Mod 2 = 0 Then b = b - 1 / a Else b = b + 1 / a
    Next i
    Label4.Caption = b
End Sub

Private Sub Command15_Click()
    Label6.Caption = "13����S= 1+1/3��+1/5��+��+1/99��"
    Dim i%, j%, a#, b#
    For i = 1 To 99 Step 2
        a = 1
        For j = 1 To i
            a = a * j
        Next j
        b = b + 1 / a
    Next i
    Label4.Caption = b
End Sub

Private Sub Command16_Click()
    Label6.Caption = "12����S=1��-3��+5��-����99��"
    Dim i%, j%, a#, b#
    For i = 1 To 99 Step 2
        a = 1
        For j = 1 To i
            a = a * j
        Next j
        If (i + 1) Mod 4 = 0 Then b = b - a Else b = b + a
    Next i
    Label4.Caption = b
End Sub

Private Sub Command17_Click()
    Label6.Caption = "11����S=1+(1+2)+(1+2+3)+����+(1+2+����+99)"
    Dim i%, j%, a#, b#
    For i = 1 To 99
        For j = 1 To i
            a = a + j
        Next j
    Next i
    Label4.Caption = a
End Sub

Private Sub Command18_Click()
    Label6.Caption = "10����1!+2!+3!+����+99!"
    Dim i%, j%, a#, b#
    For i = 1 To 99
        a = 1
        For j = 1 To i
            a = a * j
        Next j
        b = b + a
    Next i
    Label4.Caption = a
End Sub

Private Sub Command2_Click()
    Label1.Caption = "2��1-3+5-7����+99"
    Dim i%, a%
    For i = 1 To 99 Step 2
        If (i + 1) Mod 4 = 0 Then a = a - i Else a = a + i
    Next i
    Label3.Caption = a
End Sub

Private Sub Command3_Click()
    Label1.Caption = "3��1*2*3*����*99"
    Dim i%, a#
    a = 1
    For i = 1 To 99
        a = a * i
    Next i
    Label3.Caption = a
End Sub

Private Sub Command4_Click()
    Label1.Caption = "4��1*��-4��*7*��-10������*99"
    Dim i%, a%, b#
    b = 1
    For i = 1 To 99 Step 3
        a = a + 1
        If a Mod 2 = 0 Then b = b * -i Else b = b * i
    Next i
    Label3.Caption = b
End Sub

Private Sub Command5_Click()
    Label1.Caption = "5��1+3+5+����+��2N-1��"
    Dim i%, a%, n%
    n = Val(InputBox("������n��ֵ", "����", 10))
    For i = 1 To 2 * n - 1 Step 2
        a = a + i
    Next i
    Label3.Caption = a
End Sub

Private Sub Command6_Click()
    Label1.Caption = "6��2*��-4��*6*��-8������*2N"
    Dim i%, a#, b%, n#
    a = 1
    n = Val(InputBox("������n��ֵ", "����", 10))
    For i = 2 To 2 * n Step 2
        b = b + 1
        If b Mod 2 = 0 Then a = a * -i Else a = a * i
    Next i
    Label3.Caption = a
End Sub

Private Sub Command7_Click()
    Label1.Caption = "7��1+3^2+3^3+3^4+3^5+����+3^10֮������Ϣ�����"
    Dim i%, a#
    a = 1
    For i = 2 To 10
        a = a + 3 ^ i
    Next i
    Label3.Caption = a
End Sub

Private Sub Command8_Click()
    Label1.Caption = "8��ʹ��inputbox��������n�������������ǰn��֮�ͣ�1+1/3+1/5+1/7+1/9+����"
    Dim i%, a#, n%
    n = Val(InputBox("������n��ֵ", "����", 10))
    For i = 1 To 2 * n Step 2
        a = a + 1 / i
    Next i
    Label3.Caption = a
End Sub

Private Sub Command9_Click()
    Label1.Caption = "9��ʹ���������������n�������������֮����1/2*2/3*3/4*4/5*����"
    Dim i%, a#, n%
    n = Val(InputBox("������n��ֵ", "����", 10))
    a = 1
    For i = 1 To n
        a = a * i / (i + 1)
    Next i
    Label3.Caption = a
End Sub
