VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "图形打印3"
   ClientHeight    =   5685
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9495
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
   LinkTopic       =   "Form3"
   ScaleHeight     =   5685
   ScaleWidth      =   9495
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
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
      Left            =   7680
      TabIndex        =   4
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
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
      Left            =   5880
      TabIndex        =   3
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
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
      Left            =   3960
      TabIndex        =   2
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
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
      Left            =   2040
      TabIndex        =   1
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
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
      Left            =   120
      TabIndex        =   0
      Top             =   5040
      Width           =   1695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Cls
    Dim i%, j%
    Print "123456789012345678901234567890123456789012345678901234567890"
    For i = 1 To 6
        Print Tab(21 - 3 * i);
        For j = 1 To 2 * i - 1
            Print i;
        Next j
        Print
    Next i
End Sub

Private Sub Command2_Click()
    Cls
    Dim i%, j%
    Print "123456789012345678901234567890123456789012345678901234567890"
    For i = 1 To 13
        If i <= 7 Then
            Print Tab(8 - i);
            For j = 1 To 2 * i - 1
                Print "*";
            Next j
        Else
            Print Tab(i - 6);
            For j = 27 - 2 * i To 1 Step -1
                Print "*";
            Next j
        End If
        Print
    Next i
End Sub

Private Sub Command3_Click()
    Cls
    Dim i%, j%, a%, b%
    Print "123456789012345678901234567890123456789012345678901234567890"
    For i = 1 To 7
        Print Tab(8 - i);
        For j = 1 To 2 * i - 1
            Print "*";
        Next j
        Print
    Next i
    For i = 6 To 1 Step -1
        Print Tab(8 - i);
        For j = 2 * i - 1 To 1 Step -1
            Print "*";
        Next j
        Print
    Next i
End Sub

Private Sub Command4_Click()
    Cls
    Dim i%, j%
    Print "123456789012345678901234567890123456789012345678901234567890"
    For i = 1 To 13
        If i <= 7 Then
            Print Tab(8 - i);
            For j = 1 To 2 * i - 1
                If j = 1 Or j = 2 * i - 1 Then Print "&"; Else Print " ";
            Next j
        Else
            Print Tab(i - 6);
            For j = 27 - 2 * i To 1 Step -1
                If j = 1 Or j = 27 - 2 * i Then Print "&"; Else Print " ";
            Next j
        End If
        Print
    Next i
End Sub

Private Sub Command5_Click()
    Cls
    Dim i%, j%, a%, b%
    Print "123456789012345678901234567890123456789012345678901234567890"
    For i = 1 To 6
        a = i - 1
        b = -1
        Print Tab(19 - 3 * i);
        For j = 1 To 2 * i - 1
            Print Int(a);
            a = a + b
            If a = 0 Then b = -b
        Next j
        Print
    Next i
    For i = 5 To 1 Step -1
        a = i - 1
        b = -1
        Print Tab(-3 * i + 19);
        For j = 1 To 2 * i - 1
            Print Int(a);
            a = a + b
            If a = 0 Then b = -b
        Next j
        Print
    Next i
End Sub
