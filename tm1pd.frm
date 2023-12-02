VERSION 5.00
Begin VB.Form tm1pd 
   Caption         =   "题目判断"
   ClientHeight    =   10290
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16290
   LinkTopic       =   "Form1"
   ScaleHeight     =   10290
   ScaleWidth      =   16290
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "判断"
      Height          =   735
      Left            =   3480
      TabIndex        =   10
      Top             =   9360
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      Caption         =   "R2"
      Height          =   8415
      Index           =   1
      Left            =   480
      TabIndex        =   7
      Top             =   720
      Width           =   4575
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7620
         Left            =   1800
         TabIndex        =   9
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Height          =   7935
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7620
      Left            =   7200
      TabIndex        =   6
      Top             =   960
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "R1"
      Height          =   8415
      Left            =   5160
      TabIndex        =   3
      Top             =   720
      Width           =   4695
      Begin VB.TextBox text3 
         Height          =   8055
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "R2"
      Height          =   8415
      Index           =   0
      Left            =   9960
      TabIndex        =   0
      Top             =   720
      Width           =   4575
      Begin VB.TextBox Text2 
         Height          =   7935
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7620
         Left            =   1800
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "在""模式特权/en""模式下输入“show run”,从Building configuration...开始到end复制到对应的文本框中"
      Enabled         =   0   'False
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   16920
   End
End
Attribute VB_Name = "tm1pd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a() As String
Dim b As Integer
a = Split(Text1.Text, vbCrLf)
b = UBound(a)

'改名
Dim a1 As Boolean
a1 = False
For i = 0 To b
    If Trim(a(i)) <> "hostname SW2" Then
        a1 = True
    End If
Next i
    If a1 = True Then
        For i = 0 To 999
            If List1.List(i) = "" Then
            List1.List(i) = "未配置名称"
            Exit For
            End If
        Next i
    End If
If a1 = True Then
    For i = 0 To b
        If Trim(a(i)) = "hostname SW2" Then
            a1 = False
        End If
    Next i
End If
    
If a1 = False Then
    For i = 0 To 999
        If List1.List(i) = "未配置名称" Then
            List1.List(i) = ""
        End If
    Next i
End If
'interface FastEthernet0/1
'f0/1
Dim a2 As Boolean
Dim f0 As Boolean
Dim fm As Integer
Dim fm1 As Integer
f0 = False
a2 = False
For i = 0 To b
    If Trim(a(i)) = Trim("interface FastEthernet0/1") Then
        a2 = True
        fm = i
        Exit For
    End If
Next i
If a2 = True Then
    For i = fm To b
        If Trim(a(i)) = "!" Then
            fm1 = i
            Exit For
        End If
    Next i
End If
If a2 = True Then
    For i = fm To fm1
        If Trim(a(i)) <> "switchport access vlan 10" Then
            f0 = True
        End If
        If Trim(a(i)) = "switchport access vlan 10" Then
            f0 = False
        End If
    Next i
End If

If f0 = True Then
    For i = 0 To 999
        If List1.List(i) = "" Then
            List1.List(i) = "F0/1是否加入vlan"
            Exit For
        End If
    Next i
End If

If f0 = True Then
    For i = fm To fm1
        If Trim(a(i)) = "switchport access vlan 10" Then
            f0 = False
        End If
    Next i
End If

If f0 = False Then
    For i = 0 To 999
        If Trim(List1.List(i)) = "F0/1是否加入vlan" Then
            List1.List(i) = ""
        End If
    Next i
End If

'0/11
Dim a3 As Boolean
Dim f1 As Boolean
Dim fm2 As Integer
Dim fm3 As Integer
f1 = False
a3 = False
For i = 0 To b
    If Trim(a(i)) = Trim("interface FastEthernet0/11") Then
        a3 = True
        fm2 = i
        Exit For
    End If
Next i
If a3 = True Then
    For i = fm2 To b
        If Trim(a(i)) = "!" Then
            fm3 = i
            Exit For
        End If
    Next i
End If
If a3 = True Then
    For i = fm2 To fm3
        If Trim(a(i)) <> "switchport access vlan 20" Then
            f1 = True
        End If
        If Trim(a(i)) = "switchport access vlan 20" Then
            f1 = False
        End If
    Next i
End If

If f1 = True Then
    For i = 0 To 999
        If List1.List(i) = "" Then
            List1.List(i) = "F0/11是否加入vlan"
            Exit For
        End If
    Next i
End If

If f1 = True Then
    For i = fm2 To fm3
        If Trim(a(i)) = "switchport access vlan 20" Then
            f1 = False
        End If
    Next i
End If

If f1 = False Then
    For i = 0 To 999
        If Trim(List1.List(i)) = "F0/11是否加入vlan" Then
            List1.List(i) = ""
        End If
    Next i
End If
End Sub

    
