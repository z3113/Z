VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFC0C0&
   Caption         =   "style属性"
   ClientHeight    =   4575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7455
   LinkTopic       =   "Form2"
   ScaleHeight     =   4575
   ScaleWidth      =   7455
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text3 
      Height          =   855
      Left            =   5400
      TabIndex        =   11
      Top             =   3120
      Width           =   1695
   End
   Begin VB.ComboBox Combo3 
      Height          =   300
      ItemData        =   "Form2.frx":0000
      Left            =   5400
      List            =   "Form2.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   2880
      TabIndex        =   7
      Top             =   3120
      Width           =   1695
   End
   Begin VB.ComboBox Combo2 
      Height          =   1080
      ItemData        =   "Form2.frx":002C
      Left            =   2880
      List            =   "Form2.frx":0042
      Style           =   1  'Simple Combo
      TabIndex        =   6
      Text            =   "浙江"
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   360
      TabIndex        =   3
      Top             =   3120
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Form2.frx":006A
      Left            =   360
      List            =   "Form2.frx":007D
      TabIndex        =   2
      Text            =   "山东"
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "下拉组合框"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "style为2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5400
      TabIndex        =   8
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "下拉组合框"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "style为1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "下拉组合框"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "style为0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_Change()
    Text1.Text = Combo1.Text
End Sub

Private Sub Combo1_Click()
    Text1.Text = Combo1.Text
End Sub

Private Sub Combo2_Change()
    Text2.Text = Combo2.Text
End Sub

Private Sub Combo2_Click()
    Text2.Text = Combo2.Text
End Sub

Private Sub Combo3_Change()
    Text3.Text = Combo3.Text
End Sub

Private Sub Combo3_Click()
    Text3.Text = Combo3.Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
