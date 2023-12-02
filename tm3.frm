VERSION 5.00
Begin VB.Form tm3 
   Caption         =   "tm3"
   ClientHeight    =   2310
   ClientLeft      =   6870
   ClientTop       =   4470
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   ScaleHeight     =   2310
   ScaleWidth      =   11865
   Begin VB.CommandButton Command2 
      Caption         =   "题目判断"
      Height          =   855
      Left            =   8640
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "题目要求"
      Height          =   855
      Left            =   7200
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   120
      Picture         =   "tm3.frx":0000
      ScaleHeight     =   2115
      ScaleWidth      =   7035
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
   Begin VB.Label Label1 
      Caption         =   "注意:严格按照题目要求,路由器为固定组件"
      Height          =   495
      Left            =   7320
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "tm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
tmyq3bg.Show
tmyq3wz.Show
End Sub

Private Sub Command2_Click()
tm3pd.Show
End Sub
