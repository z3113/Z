VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00FFC0C0&
   Caption         =   "基本操作"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7335
   LinkTopic       =   "Form6"
   ScaleHeight     =   5595
   ScaleWidth      =   7335
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "预览"
      Height          =   495
      Left            =   4320
      TabIndex        =   17
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "重填"
      Height          =   495
      Left            =   1320
      TabIndex        =   16
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "电子竞技"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   15
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "IT技术"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   14
      Top             =   2040
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "女"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   13
      Top             =   840
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "男"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   12
      Top             =   840
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   4560
      TabIndex        =   9
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label9 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   960
      TabIndex        =   18
      Top             =   3600
      Width           =   5415
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   240
      X2              =   7080
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "性别"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "日"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   10
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "月"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "年"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "加入板块"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "出生日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "论坛昵称"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "论坛用户基本信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Check1.Value = 0
    Check2.Value = 0
End Sub

Private Sub Command2_Click()
    Dim a$, b$, c$, d$
    If Option1.Value = True Then
        a = "男"
    ElseIf Option2.Value = True Then
        a = "女"
    End If
    If Check1.Value = 1 Then b = "IT技术  " Else b = ""
    If Check2.Value = 1 Then c = "电子竞技" Else c = ""
    d = Text1.Text & vbCrLf & a & vbCrLf & Text2.Text & "年" & Text3.Text & "月" & Text4.Text & "日" & vbCrLf & b & c
    Label9.Caption = d
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
