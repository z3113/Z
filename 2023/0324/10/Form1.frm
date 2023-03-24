VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "打印图形3"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11055
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
   ScaleHeight     =   5535
   ScaleWidth      =   11055
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command7 
      Caption         =   "打印特殊菱形2"
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
      Left            =   8880
      TabIndex        =   6
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "打印特殊菱形1"
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
      Left            =   6720
      TabIndex        =   5
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "打印菱形（算法2）"
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
      Left            =   4560
      TabIndex        =   4
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "打印菱形"
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
      Left            =   2400
      TabIndex        =   3
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "打印数字三角形"
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
      Left            =   240
      TabIndex        =   2
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "复习二"
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
      Left            =   2400
      TabIndex        =   1
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "复习一"
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
      Left            =   240
      TabIndex        =   0
      Top             =   4080
      Width           =   1935
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
    Print "123456789012345678901234567890123456789012345678901234567890"
    Dim i%, j%
    For i = 1 To 5
        Print Tab(5 + i);
        For j = 1 To 5
            Print Chr(64 + i);
        Next j
    Next i
End Sub

Private Sub Command2_Click()
    Cls
    Print "123456789012345678901234567890123456789012345678901234567890"
    Dim i%, j%
    For i = 1 To 5
        Print Tab(9 - i);
        For j = 1 To 2 * i - 1
            Print Chr(64 + i);
        Next j
    Next i
End Sub

Private Sub Command3_Click()
    Cls
    Print "123456789012345678901234567890123456789012345678901234567890"
    Dim i%, j%
    For i = 1 To 6
        Print Tab(20 - 3 * i);
        For j = 1 To 2 * i - 1
            Print i;
        Next j
    Next i
End Sub

Private Sub Command4_Click()
    Cls
    Print "123456789012345678901234567890123456789012345678901234567890"
    Dim i%, j%
    For i = 1 To 7
        Print Tab(8 - i);
        For j = 1 To 2 * i - 1
            Print "*";
        Next j
    Next i
    For i = 6 To 1 Step -1
        Print Tab(8 - i);
        For j = 1 To 2 * i - 1
            Print "*";
        Next j
    Next i
End Sub

Private Sub Command5_Click()
    Cls
    Print "123456789012345678901234567890123456789012345678901234567890"
    Dim i%, j%
    For i = -6 To 6
        Print Tab(1 + Abs(i));
        For j = 1 To (13 - Abs(2 * i))
            Print "*";
        Next j
    Next i
End Sub

Private Sub Command6_Click()
    Cls
    Print "123456789012345678901234567890123456789012345678901234567890"
    Dim i%, j%
    For i = -6 To 6
        Print Tab(1 + Abs(i));
        For j = 1 To 13 - Abs(2 * i)
            If j = 1 Or j = 13 - Abs(2 * i) Then Print "&"; Else Print " ";
        Next j
    Next i
End Sub

Private Sub Command7_Click()
    Cls
    Print "123456789012345678901234567890123456789012345678901234567890"
    Dim i%, j%
    For i = -5 To 5
        Print Tab(1 + Abs(i * 3));
        For j = -Abs(5 - Abs(i)) To Abs(5 - Abs(i))
            Print Abs(j);
        Next j
    Next i
End Sub
