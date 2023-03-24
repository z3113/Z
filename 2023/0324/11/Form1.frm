VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8310
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   8310
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command10 
      Caption         =   "高考原题"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   9
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "乘法表"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   8
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "字母"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   7
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "数字沙漏"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   6
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "沙漏型"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "数字菱形3"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "数字菱形2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "字母菱形"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "数字菱形1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "数字三角形"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Cls
    Print "1234567890123456789012345678901234567890"
    Dim i%, j%
    For i = 1 To 5
        Print Tab(20 - 3 * i);
        For j = 1 To 2 * i - 1
            Print i;
        Next j
    Next i
End Sub

Private Sub Command10_Click()
    Cls
    Print "1234567890123456789012345678901234567890"
    Dim i%, j%
    For i = 1 To 10
        Print Tab(5 + i);
        For j = 5 + i To 25 + i
            If j = 26 - i Then Print "  "; Else Print "*";
        Next j
    Next i
End Sub

Private Sub Command2_Click()
    Cls
    Print "1234567890123456789012345678901234567890"
    Dim i%, j%
    For i = -5 To 5
        Print Tab(1 + 3 * Abs(i));
        For j = 1 To 11 - Abs(2 * i)
            Print 6 - Abs(i);
        Next j
    Next i
End Sub

Private Sub Command3_Click()
    Cls
    Print "1234567890123456789012345678901234567890"
    Dim i%, j%
    For i = -5 To 5
        Print Tab(10 + Abs(i));
        For j = 1 To 11 - Abs(i * 2)
            Print Chr(65 + Abs(i));
        Next j
    Next i
End Sub

Private Sub Command4_Click()
    Cls
    Print "1234567890123456789012345678901234567890"
    Dim i%, j%
    For i = -5 To 5
        Print Tab(1 + 3 * Abs(i));
        For j = Abs(i) - 5 To 5 - Abs(i)
            Print Abs(j);
        Next j
    Next i
End Sub

Private Sub Command5_Click()
    Cls
    Print "1234567890123456789012345678901234567890"
    Dim i%, j%
    For i = -5 To 5
        Print Tab(1 + 3 * Abs(i));
        For j = Abs(i) - 5 To 5 - Abs(i)
            Print 5 - Abs(j);
        Next j
    Next i
End Sub

Private Sub Command6_Click()
    Cls
    Print "1234567890123456789012345678901234567890"
    Dim i%, j%
    For i = -5 To 5
        Print Tab(6 - Abs(i));
        For j = 1 To Abs(2 * i) + 1
            Print "*";
        Next j
    Next i
End Sub

Private Sub Command7_Click()
    Cls
    Print "1234567890123456789012345678901234567890"
    Dim i%, j%
    For i = -6 To 6
        Print Tab(11 - Abs(i));
        For j = 7 - Abs(i) To 7 + Abs(i)
            If j < 10 Then Print Trim(j); Else Print Trim(Mid(j, 2, 1));
        Next j
    Next i
End Sub

Private Sub Command8_Click()
    Cls
    Print "1234567890123456789012345678901234567890"
    Dim i%, j%
    For i = 1 To 8
        Print Tab(16 - i);
        For j = 1 To 2 * i - 1
            If j >= 3 Then Print Chr(64 + j);
        Next j
    Next i
End Sub

Private Sub Command9_Click()
    Cls
    Dim i%, j%
    Print "*"; " ":
    For i = 1 To 9
        Print Trim(i); " ";
    Next i
    Print
    For i = 1 To 9
        Print Trim(i); " ";
        For j = 1 To i
            Print Trim(i * j); " ";
        Next j
        Print
    Next i
End Sub
