VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "ͼ�δ�ӡ"
   ClientHeight    =   5565
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8910
   BeginProperty Font 
      Name            =   "����"
      Size            =   14.25
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   ScaleHeight     =   5565
   ScaleWidth      =   8910
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command3 
      Caption         =   "ɳ©"
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
      Left            =   7080
      TabIndex        =   2
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
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
      Left            =   7080
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
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
      Left            =   7080
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Cls
    Print "12345678901234567890123456789012345678901234567890123456789012345678901234567890"
    Dim i%, j%, a%
    a = Val(InputBox("����������", "ͼ�δ�ӡ", 10))
    For i = 1 To a
        Print Tab(a + 1 - i);
        For j = 1 To 2 * i - 1
            Print Chr(64 + i);
        Next j
    Next i
End Sub

Private Sub Command2_Click()
    Cls
    Print "12345678901234567890123456789012345678901234567890123456789012345678901234567890"
    Dim i%, j%, a%
    a = Val(InputBox("����������(����)", "ͼ�δ�ӡ", 11))
    If a Mod 2 = 0 Then
        MsgBox "������������"
    Else
        a = (a - 1) / 2
        For i = -a To a
            Print Tab(Abs(i) + 1);
            For j = 1 To 11 - Abs(2 * i)
                If j Mod 2 = 0 Then Print "*"; Else Print "$";
            Next j
        Next i
    End If
End Sub

Private Sub Command3_Click()
    Cls
    Print "12345678901234567890123456789012345678901234567890123456789012345678901234567890"
        Dim i%, j%, a%
    a = Val(InputBox("����������(����)", "ͼ�δ�ӡ", 11))
    a = (a - 1) / 2
    For i = -a To a
        Print Tab(a - Abs(i) + 1);
        For j = 1 To Abs(2 * i) + 1
            Print Chr(65 + Abs(i));
        Next j
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
