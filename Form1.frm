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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton tmxzan 
      Caption         =   "��Ŀѡ��"
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
      Caption         =   "��ʼ"
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
      Caption         =   "��ȡ�ı�����"
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
      Caption         =   "n��ֵ��F01 F02 F03�ĸ�"
      Height          =   615
      Left            =   12960
      TabIndex        =   14
      Top             =   3600
      Width           =   2895
   End
   Begin VB.Frame Frame5 
      Caption         =   "�ı������һ�е�����"
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
Dim ���� As Integer
Dim ��������() As String

�������� = Split(Text1.Text, vbCrLf) 'Split���ڲ��
���� = UBound(��������) + 1 'UBound���ڻ�ȡ�����±�����

Text3.Text = ����
hht1 = Split(Text1.Text, vbCrLf)
Text2.Text = hht1(UBound(��������))
End Sub

Private Sub Command2_Click()
Dim a As Integer
Dim b() As String
Dim c As Integer
Dim d As String
c = 0
b = Split(Text1.Text, vbCrLf)
a = UBound(b) '��ȡ�ı�����
For i = 0 To a
    If b(i) = "���" Then Text3.Text = "�Һܺ�"
    If b(i) <> "���" Then Text3.Text = ""
Next i
Text4.Text = i

Dim ak
Dim bk
ak = 0
'kѭ�� bΪ�ı��ϼ� nΪ�̶���
'���д���������

'f01Ϊ ��� �����һ��
'f02Ϊ ��� ����ڶ���
'f03Ϊ ��� ���������
Dim f01, f02, f03 As String
Dim br As Boolean '�ú��������������ж��Ƿ��n��ֵ
Dim n
br = False
n = 0
For k = 0 To a
    If b(k) = "���" Then
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

If f01 = "���" Or f02 = "���" Or f03 = "���" Then Text5.Text = "������ܺ���"



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
'    If pd(p) <> "���" Then
'    Text5.Text = "û����ã�"
'    End If
'    If pd(p) = "���" Then
'    Text5.Text = ""
'    End If
'Next p


End Sub



Private Sub tmxzan_Click()
tmxz.Show
Unload Me
End Sub
