VERSION 5.00
Begin VB.Form tm1 
   Caption         =   "tm1"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   11055
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "题目判断"
      Height          =   735
      Left            =   9360
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "题目要求"
      Height          =   735
      Left            =   7680
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   0
      Picture         =   "tm1.frx":0000
      ScaleHeight     =   2355
      ScaleWidth      =   7515
      TabIndex        =   0
      Top             =   0
      Width           =   7575
   End
   Begin VB.Label Label1 
      Caption         =   "注意:严格按照题目要求,路由器为固定组件"
      Height          =   615
      Left            =   7680
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "tm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
tmyw1.Show
End Sub

Private Sub Command2_Click()
tm1pd.Show
End Sub
