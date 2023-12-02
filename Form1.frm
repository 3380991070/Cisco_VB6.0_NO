VERSION 5.00
Begin VB.Form cs 
   Caption         =   "cs"
   ClientHeight    =   8250
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13830
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   13830
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton tmxzan 
      Caption         =   "题目选择"
      Height          =   855
      Left            =   240
      TabIndex        =   16
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   4320
      TabIndex        =   10
      Text            =   "Text9"
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Text            =   "Text8"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Text            =   "Text7"
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   13080
      TabIndex        =   7
      Text            =   "Text6"
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   13080
      TabIndex        =   6
      Text            =   "Text5"
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   6960
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "开始"
      Height          =   735
      Left            =   11760
      TabIndex        =   4
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   13080
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   9360
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "获取文本行数"
      Height          =   735
      Left            =   11640
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   1455
      Left            =   6600
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   960
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "F01"
      Height          =   1095
      Left            =   4080
      TabIndex        =   11
      Top             =   0
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "F02"
      Height          =   1215
      Left            =   4080
      TabIndex        =   12
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      Caption         =   "F03"
      Height          =   1455
      Left            =   4080
      TabIndex        =   13
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Frame Frame4 
      Caption         =   "n的值，F01 F02 F03的根"
      Height          =   615
      Left            =   12960
      TabIndex        =   14
      Top             =   3600
      Width           =   2895
   End
   Begin VB.Frame Frame5 
      Caption         =   "文本框最后一行的内容"
      Height          =   1695
      Left            =   9120
      TabIndex        =   15
      Top             =   3480
      Width           =   2415
   End
End
Attribute VB_Name = "cs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub transparentbutton_click()
Text1.Text = Text1.Text + 1
End Sub
Private Sub Label1_Click()
Text1.Text = Text1.Text + 1
End Sub

Private Sub Command1_Click()
'Dim hht1() As String
'Dim wz As Integer
'wz = 1
'hht1 = Split(Text1.Text, vbCrLf)
'If hht1(1) = "" Then wz = 2
'Text2.Text = hht1(1)
Dim 行数 As Integer
Dim 数据数组() As String

数据数组 = Split(Text1.Text, vbCrLf) 'Split用于拆分
行数 = UBound(数据数组) + 1 'UBound用于获取数组下标上限

Text3.Text = 行数
hht1 = Split(Text1.Text, vbCrLf)
Text2.Text = hht1(UBound(数据数组))
End Sub

Private Sub Command2_Click()
Dim a As Integer
Dim b() As String
Dim c As Integer
Dim d As String
c = 0
b = Split(Text1.Text, vbCrLf)
a = UBound(b) '获取文本行数
For i = 0 To a
    If b(i) = "你好" Then Text3.Text = "我很好"
    If b(i) <> "你好" Then Text3.Text = ""
Next i
Text4.Text = i

Dim ak
Dim bk
ak = 0
'k循环 b为文本合集 n为固定数
'下列代码无问题

'f01为 你好 下面第一行
'f02为 你好 下面第二行
'f03为 你好 下面第三行
Dim f01, f02, f03 As String
Dim br As Boolean '该函数布尔型用来判断是否给n赋值
Dim n
br = False
n = 0
For k = 0 To a
    If b(k) = "你好" Then
    br = True
    Exit For
    End If
Next k
If br = True Then n = k
Text6.Text = n

f01 = ""
f02 = ""
f03 = ""
If n + 1 > a Then
    ak = bk
    ElseIf n + 1 <= a Then
      f01 = n + 1
End If

If n + 2 > a Then
    ak = bk
    ElseIf n + 2 <= a Then
      f02 = n + 2
End If

If n + 3 > a Then
    ak = bk
    ElseIf n + 3 <= a Then
      f03 = n + 3
End If
If f01 <> "" Then f01 = b(f01)
If f02 <> "" Then f02 = b(f02)
If f03 <> "" Then f03 = b(f03)
Text7.Text = f01
Text8.Text = f02
Text9.Text = f03

If f01 = "你好" Or f02 = "你好" Or f03 = "你好" Then Text5.Text = "看来你很好嘛"



'If pd(0) = "" Then
'    pd(0) = b(f01)
'    Else
'    pd(1) = b(f02)
'End If
'
'If pd(0) = "" Then
'    pd(0) = b(f03)
'    ElseIf pd(1) = "" Then
'    pd(1) = b(f03)
'    Else
'    pd(2) = b(f03)
'End If
'en = UBound(pd)
'For p = 0 To en
'    If pd(p) <> "你好" Then
'    Text5.Text = "没有你好！"
'    End If
'    If pd(p) = "你好" Then
'    Text5.Text = ""
'    End If
'Next p


End Sub



Private Sub tmxzan_Click()
tmxz.Show
Unload Me
End Sub
