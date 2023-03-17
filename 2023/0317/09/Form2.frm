VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "图形打印2"
   ClientHeight    =   4575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9015
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
   LinkTopic       =   "Form2"
   ScaleHeight     =   4575
   ScaleWidth      =   9015
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command7 
      Caption         =   "窗体3"
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
      Left            =   7560
      TabIndex        =   6
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "三角形4"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   7560
      TabIndex        =   5
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "三角形3"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6120
      TabIndex        =   4
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "平行四边形"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4680
      TabIndex        =   3
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "打印等腰三角形2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3000
      TabIndex        =   2
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "打印等腰三角形"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1320
      TabIndex        =   1
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "打印反三角形"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   1095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Cls
    Dim i%, j%
    Print "123456789012345678901234567890123456789012345678901234567890"
    For i = 1 To 5
        Print Tab(6 - i);
        For j = 1 To i
            Print "*";
        Next j
        Print
    Next i
End Sub

Private Sub Command2_Click()
    Cls
    Dim i%, j%
    Print "123456789012345678901234567890123456789012345678901234567890"
    For i = 1 To 5
        Print Tab(6 - i);
        For j = 1 To 2 * i - 1
            Print "*";
        Next j
        Print
    Next i
End Sub

Private Sub Command3_Click()
    Cls
    Dim i%, j%
    Print "123456789012345678901234567890123456789012345678901234567890"
    For i = 1 To 5
        Print Tab(6 - i);
        For j = 1 To 2 * i - 1
            If j Mod 2 = 0 Then
                Print "*";
            Else
                Print "$";
            End If
        Next j
        Print
    Next i
End Sub

Private Sub Command4_Click()
    Cls
    Dim i%, j%
    Print "123456789012345678901234567890123456789012345678901234567890"
    For i = 1 To 5
        Print Space(i - 1);
        For j = 1 To 5
            Print "@";
        Next j
        Print
    Next i
End Sub

Private Sub Command5_Click()
    Cls
    Dim i%, j%
    Print "123456789012345678901234567890123456789012345678901234567890"
    Print
    For i = 1 To 9
        Print Tab(15 - i);
        For j = 1 To i
            Print Trim(i);
        Next j
        Print
    Next i
End Sub

Private Sub Command6_Click()
    Cls
    Dim i%, j%
    Print "123456789012345678901234567890123456789012345678901234567890"
    For i = 1 To 9
        Print Tab(15 - i);
        For j = 1 To 2 * i - 1
            If j = 1 Or j = 2 * i - 1 Then
                Print "&";
            Else
                Print " ";
            End If
        Next j
        Print
    Next i
End Sub

Private Sub Command7_Click()
    Form3.Show
End Sub
