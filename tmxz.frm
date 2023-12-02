VERSION 5.00
Begin VB.Form tmxz 
   Caption         =   "tmxz"
   ClientHeight    =   6855
   ClientLeft      =   5400
   ClientTop       =   2775
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   8640
   Begin VB.CommandButton Command2 
      Caption         =   "题目1"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "题目3"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   1560
      Y1              =   1080
      Y2              =   1080
   End
End
Attribute VB_Name = "tmxz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
tm3.Show
End Sub

Private Sub Command2_Click()
tm1.Show
End Sub
