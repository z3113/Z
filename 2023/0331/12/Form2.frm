VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "������ϰ"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7095
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   6615
   ScaleWidth      =   7095
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command6 
      Caption         =   "ͼ��5"
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
      Left            =   5520
      TabIndex        =   5
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ͼ��4"
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
      Left            =   5520
      TabIndex        =   4
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ͼ��3"
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
      Left            =   5520
      TabIndex        =   3
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ͼ��2"
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
      Left            =   5520
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ͼ��1"
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
      Left            =   5520
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ˮ�ɻ���"
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
      Left            =   5520
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Form2
    Form3.Show
End Sub

Private Sub Command2_Click()
    Cls
    Print "1234567890123456789012345678901234567890"
    Dim i%, j%
    For i = 1 To 4
        Print Tab(5 - i);
        For j = 1 To 2 * i - 1
            Print Trim((j + i - 1) Mod 10);
        Next j
    Next i
End Sub

Private Sub Command3_Click()
    Cls
    Print "1234567890123456789012345678901234567890"
    Dim i%, j%
    For i = -7 To 7
        Print Tab(5 + Abs(i));
        For j = Abs(i) - 7 To 7 - Abs(i)
            Print Trim(Abs(Abs(j) - 8));
        Next j
    Next i
End Sub

Private Sub Command4_Click()
    Cls
    Print "1234567890123456789012345678901234567890"
    Dim i%, j%
    For i = -6 To 6
        Print Tab(19 - 3 * Abs(i));
        For j = -Abs(i) To Abs(i)
            Print Abs(j);
        Next j
    Next i
End Sub

Private Sub Command5_Click()
    Cls
    Print "1234567890123456789012345678901234567890"
    Dim i%, j%
    For i = -5 To 5
        Print Tab(15 - Abs(i));
        For j = -Abs(i) To Abs(i)
            If j = -Abs(i) Or j = Abs(i) Then
                If j Mod 2 = 0 Then Print "2"; Else Print "1";
            Else
                Print "*";
            End If
        Next j
    Next i
End Sub

Private Sub Command6_Click()
    Cls
    Print "1234567890123456789012345678901234567890"
    Dim i%, j%
    For i = -4 To 4
        Print Tab(5 - Abs(i));
        For j = -Abs(i) To Abs(i)
            Print Chr(65 + Abs(i) - Abs(j));
        Next j
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
