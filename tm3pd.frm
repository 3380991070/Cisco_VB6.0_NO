VERSION 5.00
Begin VB.Form tm3pd 
   Caption         =   "��Ŀ3�ж�"
   ClientHeight    =   10770
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17400
   BeginProperty Font 
      Name            =   "����"
      Size            =   18
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10770
   ScaleWidth      =   17400
   StartUpPosition =   3  '����ȱʡ
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   7680
      Max             =   3276
      Min             =   -3276
      TabIndex        =   46
      Top             =   9240
      Width           =   5055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "���"
      Height          =   735
      Left            =   10320
      TabIndex        =   45
      Top             =   9720
      Width           =   2415
   End
   Begin VB.Frame Frame6 
      Caption         =   "pc3"
      Height          =   3735
      Left            =   14520
      TabIndex        =   38
      Top             =   840
      Width           =   4935
      Begin VB.TextBox Text15 
         Height          =   495
         Left            =   1320
         TabIndex        =   41
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox Text14 
         Height          =   495
         Left            =   1320
         TabIndex        =   40
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox Text13 
         Height          =   495
         Left            =   1320
         TabIndex        =   39
         Top             =   1680
         Width           =   3375
      End
      Begin VB.Label Label17 
         Caption         =   "IP��ַ"
         Height          =   495
         Left            =   120
         TabIndex        =   44
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "��������"
         Height          =   495
         Left            =   120
         TabIndex        =   43
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "����"
         Height          =   495
         Left            =   120
         TabIndex        =   42
         Top             =   1680
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "pc4"
      Height          =   3495
      Left            =   14520
      TabIndex        =   31
      Top             =   4680
      Width           =   4815
      Begin VB.TextBox Text11 
         Height          =   495
         Left            =   1320
         TabIndex        =   34
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox Text10 
         Height          =   495
         Left            =   1320
         TabIndex        =   33
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox Text9 
         Height          =   495
         Left            =   1320
         TabIndex        =   32
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label13 
         Caption         =   "����"
         Height          =   495
         Left            =   120
         TabIndex        =   37
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "��������"
         Height          =   495
         Left            =   120
         TabIndex        =   36
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "IP��ַ"
         Height          =   495
         Left            =   120
         TabIndex        =   35
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�ж�"
      Height          =   735
      Left            =   7680
      TabIndex        =   3
      Top             =   9720
      Width           =   2535
   End
   Begin VB.Frame Frame4 
      Caption         =   "pc2"
      Height          =   3495
      Left            =   9600
      TabIndex        =   16
      Top             =   4680
      Width           =   4935
      Begin VB.CommandButton Command5 
         Caption         =   "f0/0"
         Height          =   375
         Left            =   2400
         TabIndex        =   29
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "f1/0"
         Height          =   375
         Left            =   3600
         TabIndex        =   28
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Height          =   480
         Left            =   2400
         TabIndex        =   27
         Top             =   2880
         Width           =   2295
      End
      Begin VB.TextBox pc2ip 
         Height          =   495
         Left            =   1320
         TabIndex        =   19
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox pc2ym 
         Height          =   495
         Left            =   1320
         TabIndex        =   18
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox pc2wg 
         Height          =   495
         Left            =   1320
         TabIndex        =   17
         Top             =   1680
         Width           =   3375
      End
      Begin VB.Label Label9 
         Caption         =   "��·������ӵĶ˿�"
         Height          =   855
         Left            =   240
         TabIndex        =   30
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "IP��ַ"
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "��������"
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "����"
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "R2"
      Height          =   8415
      Left            =   4920
      TabIndex        =   6
      Top             =   720
      Width           =   4575
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7620
         Left            =   1800
         TabIndex        =   8
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7815
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   7680
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   9720
      Width           =   2535
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7830
      Left            =   2160
      TabIndex        =   4
      Top             =   960
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "R1"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4695
      Begin VB.TextBox R1 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8055
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "pc1"
      Height          =   3735
      Left            =   9600
      TabIndex        =   9
      Top             =   840
      Width           =   4935
      Begin VB.TextBox Text6 
         Height          =   480
         Left            =   2280
         TabIndex        =   26
         Top             =   3000
         Width           =   2295
      End
      Begin VB.CommandButton Command3 
         Caption         =   "f1/0"
         Height          =   375
         Left            =   3480
         TabIndex        =   25
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "f0/0"
         Height          =   375
         Left            =   2280
         TabIndex        =   24
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   1320
         TabIndex        =   15
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   1320
         TabIndex        =   14
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   1320
         TabIndex        =   10
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label8 
         Caption         =   "��·������ӵĶ˿�"
         Height          =   855
         Left            =   120
         TabIndex        =   23
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "����"
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "��������"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "IP��ַ"
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��""ģʽ��Ȩ/en""ģʽ�����롰show run��,��Building configuration...��ʼ��end���Ƶ���Ӧ���ı�����"
      Enabled         =   0   'False
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   16920
   End
End
Attribute VB_Name = "tm3pd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim f1 As Integer
Dim f2 As Integer
Dim f3 As Integer
Dim f4 As Integer
Dim f5 As Integer
Dim f6 As Integer
Dim l1 As Integer
Private Sub Command1_Click()
'R1����
Dim R1pas '����
Dim R1pasjs As Boolean '������������ж�
R1pasjs = False
Dim R1pasjs1 '�������

Dim R1host
Dim R1hostjs As Boolean
R1hostjs = False
Dim R1hostjs1
R1hostjs1 = -1

Dim R1sl '�ı������ж�
Dim R1lj
Dim R1pd
Dim R1pd1

R1lj = Split(R1.Text, vbCrLf)
R1sl = UBound(R1lj) '��ȡ�ı�����
'R1pd = Split(R1zd.Text, vbCrLf)
'R1pd1 = UBound(R1pd)
    '���д������������ж�R1host
    For i = 0 To R1sl
        If R1lj(i) = "hostname Router" Then
            R1hostjs = True
            R1hostjs1 = i
            Exit For
            ElseIf R1lj(i) = "hostname R1" Then
            R1hostjs = True
            R1hostjs1 = i
            Exit For
        End If
        'Or R1lj(i) = "hostname R1" Then
        '    R1hostjs = True
        '    Exit For
        'End If
    Next i
    If R1hostjs = True Then R1hostjs1 = i
    '�������
    'Text2.Text = I
    
    Dim R1hostjs2 As Boolean
    Dim R1k
    Dim R1c As Boolean
    R1c = False
    If R1hostjs1 >= 0 Then R1c = True
    
If R1c = True Then
    If R1lj(i) = "hostname Router" Then
        R1hostjs2 = False
            ElseIf R1lj(i) = "hostname R1" Then
                R1hostjs2 = True
                    Else
                        R1k = 1
                        End If
            End If
            'If List1.List(0) = "" Then List1.List(0) = "���"
If R1hostjs2 = False Then
    List1.List(0) = "δ��������"
    ElseIf R1hostjs2 = True Then
    R1k = 2
    ElseIf R1k = 1 Then
    List1.List(0) = "δ��ȷ��������"
End If
If R1k = 2 Then List1.List(0) = ""
'If R1hostjs2 = False Then
'        R1zd.Text = "δ��������"
'        ElseIf R1hostjs2 = True Then
'        R1k = 2
'        ElseIf R1k = 1 Then
'        R1zd.Text = "δ��������"
'End If

'If R1k = 2 Then R1zd.Text = ""
     
'If R1pd1 = -1 Then R1zd.Text = 1
'If R1pd(0) = "1" Then R1zd.Text = "���"
'If R1pd1 = -1 Then R1zd.Text = vbCrLf
' For a = 0 To R1pd1
'        If R1pd(a) = "" And R1hostjs2 = False Then
'            R1pd(a) = "δ��������"
'                ElseIf R1pd(a) = "" And R1hostjs2 = True Then
'                    R1k = 2
'                        ElseIf R1pd(a) = "" And R1k = 1 Then
'                            R1pd(a) = "δ��ȷ��������"
'                                Else
'                                    R1k = 0
'        Exit For
'        End If
'    Next a

'If R1pd1 = -1 Then
'    R1zd.Text = vbCrLf
'Else
'    For a = 0 To R1pd1
'        If R1pd(a) = "" Then
'            If R1hostjs2 = False Then
'                R1pd(a) = "δ��������"
'            ElseIf R1hostjs2 = True Then
'                R1k = 2
'            ElseIf R1k = 1 Then
'                R1pd(a) = "δ��ȷ��������"
'            End If
'            Exit For
'        End If
'    Next a
'End If
    'Text1.Text = R1pd1

'����Ϊ�ж� ·���������Ƿ���ȷ
For a = 0 To R1sl
    If R1lj(a) <> "enable password 123456" Then
    R1pasjs = True
'    Text4.Text = a
    'Exit For
    ElseIf R1lj(a) = "enable password 123456" Then
    R1pasjs = False
'    Text4.Text = a
    Exit For '�������ж���ɺ�ֹͣ��������������жϽ�û����Խ����ɴ��󣡣�������ѵ��
    End If
'Exit For
Next a
    If R1pasjs = True Then
        For b = 0 To 99
        'If List1.List(b) <> "�Ƿ���ȷ����enable������" Then
            If List1.List(b) = "" Then
            List1.List(b) = "�Ƿ���ȷ����enable������"
            'Text1.Text = b
            Exit For
            End If
        Next b
    End If
    
For d = 0 To R1sl
    If R1lj(d) = "enable password 123456" Then R1pasjs = False
Exit For
Next d

    If R1pasjs = False Then
        For c = 0 To 99
            If List1.List(c) = "�Ƿ���ȷ����enable������" Then
            List1.List(c) = ""
            End If
        Next c
    End If
    
If R1pasjs = False Then
    If List1.List(0) = "�Ƿ���ȷ����enable������" Then List1.List(0) = ""
End If
'Text3.Text = R1pasjs

'����Ϊ�ж�enable������ �������������е��߼����� ƿ����
Dim R1user
Dim R1user1 As Boolean
R1user1 = False
For e = 0 To R1sl
    If R1lj(e) <> "username R2 password 0 jncs" Then
    R1user1 = True
    ElseIf R1lj(e) = "username R2 password 0 jncs" Then
    R1user1 = False
    Exit For
    End If
Next e

If R1user1 = True Then
    For f = 0 To 99
        If List1.List(f) = "" Then
        List1.List(f) = "�Ƿ���ȷ������֤��"
        Exit For
        End If
    Next f
End If

If R1user1 = False Then
    For g = 0 To 99
    If List1.List(g) = "�Ƿ���ȷ������֤��" Then List1.List(g) = ""
    Next g
End If
'���������������������߼����⣿��û�й���ʵ���޷��жϣ�
'�������������ж���֤username R2 password 0 jncs
'�������� for - g
'Aa  Bb  Cc  Dd  Ee  Ff  Gg  Hh  Ii  Jj  Kk  Ll  Mm  Nn  Oo  Pp  Qq  Rr  Ss  Tt  Uu  Vv  Ww  Xx  Yy  Zz

Dim R1f00 'F0/0����
Dim R1f00js As Boolean '���������ж�
R1f00js = False
Dim R1f00js1
Dim R1f00js2 '��h
Dim R1f00js3 '��j
R1f00js2 = -1
R1f00js3 = -1
R1f00js4 = -1 '����
For h = 0 To R1sl
    If Trim(R1lj(h)) = Trim("interface FastEthernet0/0") Then
    R1f00js2 = h '���
    R1f00js = True
        'For j = 0 To 8
        '    If R1f00js2 + j > R1sl - 1 Then
        '    R1f00js3 = j
        '    R1f00js4 = R1f00js2 + R1f00js3
    Exit For
    End If
        'Next j
    'End If
Next h
If R1f00js = True Then
    For a = R1f00js2 To R1sl
        If Trim(R1lj(a)) <> Trim("!") Then
            R1f00js3 = a
        End If
        If Trim(R1lj(a)) = Trim("!") Then
            R1f00js3 = a
            Exit For
        End If
    Next a
End If
'Text1.Text = R1f00js2 & " " & R1f00js3
'''''
'������
'Text6.Text = R1f00js2
'Text7.Text = R1f00js3 & " , " & R1f00js
'�� ������

Dim R1f00ip As Boolean '�ж�IP�Ƿ���ȷ
Dim R1f00st As Boolean '�ж϶˿��Ƿ��
R1f00st = False
R1f00ip = False
If R1f00js = True Then
    For k = R1f00js2 To R1f00js3
        'If Trim(R1lj(k)) <> Trim("!") Then
            If Trim(R1lj(k)) <> Trim("ip address 192.168.1.30 255.255.255.240") Then
            R1f00ip = True
            End If
    Next k
End If
If R1f00js = True Then
    For k = R1f00js2 To R1f00js3
            If Trim(R1lj(k)) = Trim("ip address 192.168.1.30 255.255.255.240") Then
            R1f00ip = False
            End If
        Next k
End If
If R1f00js = True Then
    For k = R1f00js2 To R1f00js3
            If Trim(R1lj(k)) <> Trim("shutdown") Then
            R1f00st = False
            End If
    Next k
End If
If R1f00js = True Then
    For k = R1f00js2 To R1f00js3
            If Trim(R1lj(k)) = Trim("shutdown") Then
            R1f00st = True
            End If
    Next k
End If


If R1f00ip = True Then
    For l = 0 To 999
        If List1.List(l) = "" Then
            List1.List(l) = "F0/0���Ƿ���ȷ����IP��ַ"
        Exit For
        End If
    Next l
End If

If R1f00st = True Then
    For m = 0 To 999
        If List1.List(m) = "" Then
            List1.List(m) = "F0/0�ж˿��Ƿ�����"
        Exit For
        End If
    Next m
End If

If R1f00st = False Then
    For m = 0 To 999
        If List1.List(m) = "F0/0�ж˿��Ƿ�����" Then
            List1.List(m) = ""
        End If
    Next m
 End If
 
 If R1f00ip = False Then
    For m = 0 To 999
        If List1.List(m) = "F0/0���Ƿ���ȷ����IP��ַ" Then
            List1.List(m) = ""
        End If
    Next m
End If

        
'������
'Text5.Text = R1f00ip

'���������ж�f0/0

Dim R1f10 'F1/0����
Dim R1f10s As Boolean
R1f10js = False
Dim R1f10js1 '�洢 Ѱ��
Dim R1f10js2 '�洢 �յ�
Dim R1f10js3 '�洢 ѭ��
R1f10js1 = -1
R1f10js2 = -1
For a = 0 To R1sl
    If Trim(R1lj(a)) = Trim("interface FastEthernet1/0") Then
        R1f10js = True
        R1f10js1 = a
        Exit For
    End If
Next a

If R1f10js = True Then
            For a = 0 To 8
                If R1f10js1 + a < R1sl Then
                If Trim(R1lj(R1f10js1 + a)) <> Trim("!") Then
                    R1f10js2 = R1f10js1 + a
                    If Trim(R1lj(R1f10js1 + a)) = Trim("!") Then
                        R1f10js2 = R1f10js1 + a
                     Exit For
                    End If
                    End If
                End If
            Next a
End If
'Text8.Text = R1f10js2 & R1f10js & " " & R1f10js1
'������

Dim R1f10ip As Boolean
Dim R1f10st As Boolean
R1f10ip = False
R1f10st = False
If R1f10js = True Then
    For a = R1f10js1 To R1f10js2
        'If Trim(R1lj(a)) <> Trim("!") Then
        If Trim(R1lj(a)) <> Trim("ip address 192.168.1.46 255.255.255.240") Then
        R1f10ip = True
        End If
    Next a
End If
If R1f10js = True Then
    For a = R1f10js1 To R1f10js2
    If Trim(R1lj(a)) <> Trim("shutdown") Then
    R1f10st = False
    End If
    Next a
End If
If R1f10js = True Then
    For a = R1f10js1 To R1f10js2
        'If Trim(R1lj(a)) <> Trim("!") Then
        If Trim(R1lj(a)) = Trim("ip address 192.168.1.46 255.255.255.240") Then
        R1f10ip = False
        End If
    Next a
End If
If R1f10js = True Then
    For a = R1f10js1 To R1f10js2
    If Trim(R1lj(a)) = Trim("shutdown") Then
    R1f10st = True
    End If
    Next a
End If
'Text8.Text = R1f10ip & " " & R1f10st
'�ж�IP�����
If R1f10ip = True Then
    For a = 0 To 999
        If List1.List(a) = "" Then
        List1.List(a) = "F1/0��IP�Ƿ���ȷ���ã�"
        Exit For
        End If
    Next a
End If
'�ж϶˿ڲ����
If R1f10st = True Then
    For a = 0 To 999
        If List1.List(a) = "" Then
            List1.List(a) = "F1/0�ж˿��Ƿ�����"
            Exit For
        End If
    Next a
End If
'��ֹ����ip
If R1f10ip = False Then
    For a = 0 To 999
        If List1.List(a) = "F1/0��IP�Ƿ���ȷ���ã�" Then List1.List(a) = ""
    Next a
End If
'��ֹ���ж˿�
If R1f10st = False Then
    For a = 0 To 999
        If List1.List(a) = "F1/0�ж˿��Ƿ�����" Then List1.List(a) = ""
    Next a
End If
'����Ϊ�ж�F1/0 ������������ û�й���ʵ�飩

    
Dim R1s20 'S2/0����
Dim R1s20js As Boolean
R1s20js = False
Dim R1s20js1 '���� ����
Dim R1s20js2 '���� β��
R1s20js1 = -1
R1s20js2 = -1

'�ж�S2/0
For a = 0 To R1sl
    If Trim(R1lj(a)) = "interface Serial2/0" Then
        R1s20js = True
        R1s20js1 = a 'S2/0����
        Exit For
    End If
Next a
'�жϽ�β
If R1s20js = True Then
    For a = R1s20js1 To R1sl
        If a < R1sl Then
            If Trim(R1lj(a)) = "!" Then
                R1s20js2 = a 'S2/0��β
                Exit For
            If Trim(R1lj(a)) <> Trim("!") Then
                R1s20js2 = a
            End If
            End If
        End If
    Next a
End If
'��֤
'ip;en ppp;ppp au chap;ʱ��
Dim R1s20ip As Boolean
Dim R1s20en As Boolean
Dim R1s20au As Boolean
Dim R1s20sz As Boolean
Dim R1s20st As Boolean
R1s20ip = False
R1s20en = False
R1s20au = False
R1s20sz = False
R1s20st = False

If R1s20js = True Then
    For a = R1s20js1 To R1s20js2
        If Trim(R1lj(a)) <> Trim("ip address 12.0.0.1 255.0.0.0") Then
        R1s20ip = True
        End If
    Next a
End If
If R1s20js = True Then
    For a = R1s20js1 To R1s20js2
        If Trim(R1lj(a)) = Trim("ip address 12.0.0.1 255.0.0.0") Then
        R1s20ip = False
        End If
    Next a
End If
If R1s20js = True Then
    For a = R1s20js1 To R1s20js2
        If Trim(R1lj(a)) <> Trim("encapsulation ppp") Then '��װpppЭ��
        R1s20en = True
        End If
    Next a
End If
If R1s20js = True Then
    For a = R1s20js1 To R1s20js2
        If Trim(R1lj(a)) = Trim("encapsulation ppp") Then '��װpppЭ��
        R1s20en = False
        End If
    Next a
End If
If R1s20js = True Then
    For a = R1s20js1 To R1s20js2
        If Trim(R1lj(a)) <> Trim("ppp authentication chap") Then '����chap��֤
        R1s20au = True
        End If
    Next a
End If
If R1s20js = True Then
    For a = R1s20js1 To R1s20js2
        If Trim(R1lj(a)) = Trim("ppp authentication chap") Then '����chap��֤
        R1s20au = False
        End If
    Next a
End If
If R1s20js = True Then
    For a = R1s20js1 To R1s20js2
        If Trim(R1lj(a)) <> Trim("clock rate 64000") Then 'ʱ��Ƶ��
        R1s20sz = True
        End If
    Next a
End If
If R1s20js = True Then
    For a = R1s20js1 To R1s20js2
        If Trim(R1lj(a)) = Trim("clock rate 64000") Then 'ʱ��Ƶ��
        R1s20sz = False
        End If
    Next a
End If
If R1s20js = True Then
    For a = R1s20js1 To R1s20js2
        If Trim(R1lj(a)) = Trim("shutdown") Then '�˿ڿ���
        R1s20st = True
        End If
    Next a
End If
If R1s20js = True Then
    For a = R1s20js1 To R1s20js2
        If Trim(R1lj(a)) <> Trim("shutdown") Then '�˿ڿ���
        R1s20st = False
        End If
    Next a
End If

'Text1.Text = R1s20ip & " " & R1s20en & " " & R1s20au & " " & R1s20sz & " " & R1s20st

If R1s20ip = True Then
    For a = 0 To 999
        If List1.List(a) = "" Then
            List1.List(a) = "S2/0�Ƿ���ȷ����IP��"
            Exit For
        End If
    Next a
End If
If R1s20ip = False Then
    For a = 0 To 999
        If List1.List(a) = Trim("S2/0�Ƿ���ȷ����IP��") Then
            List1.List(a) = ""
        End If
    Next a
End If

If R1s20en = True Then
    For a = 0 To 999
     If List1.List(a) = "" Then
        List1.List(a) = "S2/0�Ƿ��װpppЭ�飿"
        Exit For
    End If
    Next a
End If
If R1s20en = False Then
    For a = 0 To 999
        If List1.List(a) = Trim("S2/0�Ƿ��װpppЭ�飿") Then
            List1.List(a) = ""
        End If
    Next a
End If

If R1s20au = True Then
    For a = 0 To 999
        If List1.List(a) = "" Then
            List1.List(a) = "S2/0�Ƿ���chap��֤��"
            Exit For
        End If
    Next a
End If
If R1s20au = False Then
    For a = 0 To 999
        If List1.List(a) = "S2/0�Ƿ���chap��֤��" Then
            List1.List(a) = ""
        End If
    Next a
End If

If R1s20sz = True Then
    For a = 0 To 999
        If List1.List(a) = "" Then
            List1.List(a) = "S2/0�Ƿ���ȷ����ʱ��Ƶ�ʣ�"
            Exit For
        End If
    Next a
End If
If R1s20sz = False Then
    For a = 0 To 999
        If List1.List(a) = Trim("S2/0�Ƿ���ȷ����ʱ��Ƶ�ʣ�") Then
            List1.List(a) = ""
        End If
    Next a
End If

If R1s20st = True Then
    For a = 0 To 999
        If List1.List(a) = "" Then
            List1.List(a) = "S2/0�˿��Ƿ�����"
            Exit For
        End If
    Next a
End If
If R1s20st = False Then
    For a = 0 To 999
        If List1.List(a) = Trim("S2/0�˿��Ƿ�����") Then
            List1.List(a) = ""
        End If
    Next a
End If
'Text8.Text = R1s20js1 & " " & " " & R1s20js & R1s20js2

'����Ϊ�ж�s2/0

Dim R1route '��̬·������
Dim R1route1
Dim R1routejs As Boolean
R1routejs = False
Dim R1routejs1 As Boolean
R1routejs1 = False
For a = 0 To R1sl
    If Trim(R1lj(a)) = Trim("ip classless") Then
        R1routejs = True
        R1route = a
        Exit For
    End If
Next a
If R1routejs = True Then
    For a = R1route To R1sl
        If Trim(R1lj(a)) = Trim("!") Then
           R1route1 = a
            Exit For
        End If
            If Trim(R1lj(a)) <> Trim("!") Then
                R1route1 = a
            End If
    Next a
End If
If R1routejs = True Then
    For a = R1route To R1route1
        If Trim(R1lj(a)) <> Trim("ip route 192.168.2.0 255.255.255.0 12.0.0.2") Then
            R1routejs1 = True
        End If
    Next a
End If
If R1routejs = True Then
    For a = R1route To R1route1
        If Trim(R1lj(a)) = Trim("ip route 192.168.2.0 255.255.255.0 12.0.0.2") Then
            R1routejs1 = False
        End If
    Next a
End If
If R1routejs1 = True Then
    For a = 0 To 999
        If List1.List(a) = "" Then
            List1.List(a) = Trim("R1���Ƿ���ȷ����·��.")
        Exit For
        End If
    Next a
End If
If R1routejs1 = False Then
    For a = 0 To 999
        If List1.List(a) = Trim("R1���Ƿ���ȷ����·��.") Then
            List1.List(a) = ""
        End If
    Next a
End If
'Text1.Text = R1route & "  " & R1route1
    
Dim R1line 'Զ�̵�����������
Dim R1linejs As Boolean
R1linejs = False
Dim R1linejs1
For a = 0 To R1sl
    If Trim(R1lj(a)) = Trim("line vty 0 4") Then
        R1linejs = True
        R1line = a
        Exit For
    End If
Next a
If R1linejs = True Then
    For a = R1line To R1sl
        If Trim(R1lj(a)) = "!" Then
            R1linejs1 = a
            Exit For
        End If
        If Trim(R1lj(a)) <> "!" Then
            R1linejs1 = a
        End If
    Next a
End If
Dim R1linejs2 As Boolean
R1linejs2 = False
If R1linejs = True Then
    For a = R1line To R1linejs1
        If Trim(R1lj(a)) <> Trim("password 123456") Then
            R1linejs2 = True
            Else
                R1linejs2 = False
        End If
    Next a
End If
If R1linejs2 = True Then
    For a = 0 To 999
        If List1.List(a) = "" Then
            List1.List(a) = "�Ƿ���ȷ����Զ�̵�������"
        Exit For
        End If
    Next a
End If
If R1linejs2 = False Then
    For a = 0 To 999
        If Trim(List1.List(a)) = Trim("�Ƿ���ȷ����Զ�̵�������") Then
            List1.List(a) = ""
        End If
    Next a
End If
If R1linejs = False Then
    For a = 0 To 999
        If Trim(List1.List(a)) = "" Then
            List1.List(a) = Trim("�Ƿ�Ҫ����Զ�̵���")
        Exit For
        End If
    Next a
End If
If R1linejs = True Then
    For a = 0 To 999
        If Trim(List1.List(a)) = Trim("�Ƿ�Ҫ����Զ�̵���") Then
            List1.List(a) = ""
        End If
    Next a
End If

'����ΪR2
Dim R2
Dim R2a
R2 = Split(Text2.Text, vbCrLf)
R2a = UBound(R2) '��ȡ�ı�����
'�ж�·��������
Dim R2h As Boolean
R2h = False
For i = 0 To R2a
    If Trim(R2(i)) = Trim("hostname R2") Then
        R2h = True
        Exit For
    End If
Next i
If R2h = False Then
    For i = 0 To 999
        If List2.List(i) = "" Then
            List2.List(i) = "R2�У������Ƿ���ȷ"
            Exit For
        End If
    Next i
End If
If R2h = True Then
    For i = 0 To 999
        If List2.List(i) = "R2�У������Ƿ���ȷ" Then
            List2.List(i) = ""
        End If
    Next i
End If

'����Ϊ�ж���֤
Dim R2u As Boolean
R2u = False
For i = 0 To R2a
    If Trim(R2(i)) = Trim("username R1 password 0 jncs") Then
        R2u = True
        Exit For
    End If
Next i
If R2u = False Then
    For i = 0 To 999
        If List2.List(i) = "" Then
            List2.List(i) = "R2�У��Ƿ���ȷ������֤"
            Exit For
        End If
    Next i
End If
If R2u = True Then
    For i = 0 To 999
        If List2.List(i) = "R2�У��Ƿ���ȷ������֤" Then
            List2.List(i) = ""
        End If
    Next i
End If
'����Ϊ�ж�F0/0
Dim R20 As Boolean
Dim R201 'ͷ
Dim R202 'β
R20 = False
For i = 0 To R2a
    If Trim(R2(i)) = Trim("interface FastEthernet0/0") Then
        R20 = True
        R201 = i    'ͷ
        Exit For
    End If
Next i

If R20 = True Then
    For i = 0 To R2a
        If Trim(R2(i)) = Trim("!") Then
            R202 = i
            Exit For
        End If
        If Trim(R2(i)) <> Trim("!") Then
            R202 = i
        End If
    Next i
End If
Dim R20i As Boolean 'f0/0 ip
Dim R20s As Boolean 'f0/0 ����
R20i = False
R20s = False
If R20 = True Then
    For i = R201 To R202
        If Trim(R2(i)) <> "ip address 192.168.2.254 255.255.255.0" Then
        R20i = True
            Else
                R20i = False
        End If
        If Trim(R2(i)) = "shutdown" Then
        R20s = True
            Else
                R20s = False
        End If
    Next i
End If
If R20i = True Then
    For i = 0 To 999
        If List2.List(i) = "" Then
            List2.List(i) = "R2��F0/0�Ƿ���ȷ����ip��ַ"
            Exit For
        End If
    Next i
End If
If R20i = False Then
    For i = 0 To 999
        If List2.List(i) = "R2��F0/0�Ƿ���ȷ����ip��ַ" Then
            List2.List(i) = ""
        End If
    Next i
End If
If R20s = True Then
    For i = 0 To 999
        If List2.List(i) = "" Then
            List2.List(i) = "R2��F0/0�Ƿ���"
            Exit For
        End If
    Next i
End If
If R20s = False Then
    For i = 0 To 999
        If List2.List(i) = "R2��F0/0�Ƿ���" Then
            List2.List(i) = ""
        End If
    Next i
End If
'S2/0
Dim R2S2c '��ʼ
Dim R2S2j '��ֹ
Dim R2S2 As Boolean
Dim R2S2i As Boolean    'IP
Dim R2S2s As Boolean    '����
Dim R2S2e As Boolean    '��װppp
Dim R2S2p As Boolean 'chap��֤
R2S2i = False
R2S2s = False
R2S2e = False
R2S2p = False
R2S2 = False
For a = 0 To R2a
    If Trim(R2(a)) = "interface Serial2/0" Then
    R2S2 = True
    R2S2c = a
    Exit For
    End If
Next a
If R2S2 = True Then
    For a = R2S2c To R2a
        If Trim(R2(a)) = "!" Then
            R2S2j = a
            Exit For
        End If
        If Trim(R2(a)) <> "!" Then
            R2S2j = a
        End If
    Next a
End If
If R2S2 = True Then
    For a = R2S2c To R2S2j
        If Trim(R2(a)) <> "ip address 12.0.0.2 255.0.0.0" Then R2S2i = True
    Next a
End If
If R2S2 = True Then
    For a = R2S2c To R2S2j
        If Trim(R2(a)) = "ip address 12.0.0.2 255.0.0.0" Then R2S2i = False
    Next a
End If
If R2S2 = True Then
    For a = R2S2c To R2S2j
        If Trim(R2(a)) <> "encapsulation ppp" Then R2S2e = True
    Next a
End If
If R2S2 = True Then
    For a = R2S2c To R2S2j
        If Trim(R2(a)) = "encapsulation ppp" Then R2S2e = False
    Next a
End If
If R2S2 = True Then
    For a = R2S2c To R2S2j
        If Trim(R2(a)) <> "ppp authentication chap" Then R2S2p = True
    Next a
End If
If R2S2 = True Then
    For a = R2S2c To R2S2j
        If Trim(R2(a)) = "ppp authentication chap" Then R2S2p = False
    Next a
End If
If R2S2 = True Then
    For a = R2S2c To R2S2j
        If Trim(R2(a)) <> "shutdown" Then R2S2s = False
    Next a
End If
If R2S2 = True Then
    For a = R2S2c To R2S2j
        If Trim(R2(a)) = "shutdown" Then R2S2s = True
    Next a
End If
If R2S2i = True Then
    For a = 0 To 999
        If List2.List(a) = "" Then
            List2.List(a) = "R2��S2/0��IP�Ƿ���ȷ����"
        Exit For
        End If
    Next a
End If
If R2S2i = False Then
    For a = 0 To 999
        If List2.List(a) = "R2��S2/0��IP�Ƿ���ȷ����" Then
            List2.List(a) = ""
        End If
    Next a
End If
If R2S2e = True Then
    For a = 0 To 999
        If List2.List(a) = "" Then
            List2.List(a) = "R2��S2/0�У��Ƿ�����ppp"
            Exit For
        End If
    Next a
End If
If R2S2e = False Then
    For a = 0 To 999
        If List2.List(a) = "R2��S2/0�У��Ƿ�����ppp" Then
            List2.List(a) = ""
        End If
    Next a
End If
If R2S2p = True Then
    For a = 0 To 999
        If List2.List(a) = "" Then
            List2.List(a) = "R2��S2/0�У��Ƿ���chap��֤"
            Exit For
        End If
    Next a
End If
If R2S2p = False Then
    For a = 0 To 999
        If List2.List(a) = "R2��S2/0�У��Ƿ���chap��֤" Then
            List2.List(a) = ""
        End If
    Next a
End If
If R2S2s = True Then
    For a = 0 To 999
        If List2.List(a) = "" Then
            List2.List(a) = "R2��S2/0�˿��Ƿ���"
            Exit For
        End If
    Next a
End If
If R2S2s = False Then
    For a = 0 To 999
        If List2.List(a) = "R2��S2/0�˿��Ƿ���" Then
            List2.List(a) = ""
        End If
    Next a
End If

'����Ϊ·��
Dim R2r 'ͷ
Dim R2r1 'β
Dim R2rr As Boolean
Dim R2ri1 As Boolean '��һ
Dim R2ri2 As Boolean '�ڶ�
R2ri1 = False
R2ri2 = False
R2rr = False
For a = 0 To R2a
    If Trim(R2(a)) = "ip classless" Then
        R2rr = True
        R2r = a
        Exit For
    End If
Next a
If R2rr = True Then
    For a = R2r To R2a
        If Trim(R2(a)) = "!" Then
            R2r1 = a
            Exit For
        End If
        If Trim(R2(a)) <> "!" Then
            R2r1 = a
        End If
    Next a
End If
If R2rr = True Then
    For a = R2r To R2r1
        If Trim(R2(a)) <> "ip route 192.168.1.16 255.255.255.240 12.0.0.1" Then
            R2ri1 = True
        End If
    Next a
End If
If R2rr = True Then
    For a = R2r To R2r1
        If Trim(R2(a)) = "ip route 192.168.1.16 255.255.255.240 12.0.0.1" Then
            R2ri1 = False
        End If
    Next a
End If
If R2rr = True Then
    For a = R2r To R2r1
        If Trim(R2(a)) <> Trim("ip route 192.168.1.32 255.255.255.240 12.0.0.1") Then
            R2ri2 = True
        End If
    Next a
End If
If R2rr = True Then
    For a = R2r To R2r1
        If Trim(R2(a)) = Trim("ip route 192.168.1.32 255.255.255.240 12.0.0.1") Then
            R2ri2 = False
        End If
    Next a
End If

If R2ri1 = True Then
    For a = 0 To 999
        If List2.List(a) = "" Then
            List2.List(a) = "16·���Ƿ�����"
            Exit For
        End If
    Next a
End If
If R2ri1 = False Then
    For a = 0 To 999
        If List2.List(a) = "16·���Ƿ�����" Then
            List2.List(a) = ""
        End If
    Next a
End If
If R2ri2 = True Then
    For a = 0 To 999
        If List2.List(a) = "" Then
            List2.List(a) = "32·���Ƿ�����"
            Exit For
        End If
    Next a
End If
If R2ri2 = False Then
    For a = 0 To 999
        If List2.List(a) = "32·���Ƿ�����" Then
            List2.List(a) = ""
        End If
    Next a
End If
Text1.Text = R2ri2 & " " & R2r1 & " " & R2r
'PC1
'text3 = IP��ַ
'text4 = ��������
'text5 = ����
If Trim(Text3.Text) = "" Then Text3.Text = "δ��д"
If Trim(Text4.Text) = "" Then Text4.Text = "δ��д"
If Trim(Text5.Text) = "" Then Text5.Text = "δ��д"
If Trim(Text3.Text) <> Trim("192.168.1.17") And Trim(Text3.Text) <> "δ��д" Then
    Text3.Text = "����"
End If
If Trim(Text4.Text) <> Trim("255.255.255.240") And Trim(Text4.Text) <> "δ��д" Then
    Text4.Text = "����"
End If
If Trim(Text5.Text) <> Trim("192.168.1.30") And Trim(Text5.Text) <> "δ��д" Then
    Text5.Text = "����"
End If
If Trim(Text6.Text) = "" Then
    Text6.Text = "δѡ��"
End If
If Trim(Text6.Text) <> "" And Trim(Text6.Text) <> "δѡ��" And Trim(Text6.Text) <> "f0/0" Then
    Text6.Text = "����"
End If
    
            
'pc2
'pc2ip
'pc2ym
'pc2wg
If Trim(pc2ip.Text) = "" Then pc2ip.Text = "δ����"
If Trim(pc2ym.Text) = "" Then pc2ym.Text = "δ����"
If Trim(pc2wg.Text) = "" Then pc2wg.Text = "δ����"
If Trim(pc2ip.Text) <> Trim("192.168.1.33") And Trim(pc2ip.Text) <> "δ����" Then
    pc2ip.Text = "����"
End If
If Trim(pc2ym.Text) <> Trim("255.255.255.240") And Trim(pc2ym.Text) <> "δ����" Then
    pc2ym.Text = "����"
End If
If Trim(pc2wg.Text) <> Trim("192.168.1.46") And Trim(pc2wg.Text) <> "δ����" Then
    pc2wg.Text = "����"
End If
If Trim(Text7.Text) = "" Then
    Text7.Text = "δѡ��"
End If
If Trim(Text7.Text) <> "" And Trim(Text7.Text) <> "δѡ��" And Trim(Text7.Text) <> "f1/0" Then
    Text7.Text = "����"
End If

'pc3
'text13 - 15
'���� / ���� / IP
If Trim(Text13.Text) = "" Then Text13.Text = "δ����"
If Trim(Text14.Text) = "" Then Text14.Text = "δ����"
If Trim(Text15.Text) = "" Then Text15.Text = "δ����"
If Trim(Text13.Text) <> Trim("192.168.2.254") And Trim(Text13.Text) <> "δ����" Then
    Text13.Text = "����"
End If
If Trim(Text14.Text) <> Trim("255.255.255.0") And Trim(Text14.Text) <> "δ����" Then
    Text14.Text = "����"
End If
If Trim(Text15.Text) <> Trim("192.168.2.1") And Trim(Text15.Text) <> "δ����" Then
    Text15.Text = "����"
End If
'pc4
'text9 - 11
'ip / ���� / ����
If Trim(Text9.Text) = "" Then Text9.Text = "δ��д"
If Trim(Text10.Text) = "" Then Text10.Text = "δ��д"
If Trim(Text11.Text) = "" Then Text11.Text = "δ��д"
If Trim(Text9.Text) <> Trim("192.168.2.2") And Trim(Text9.Text) <> "δ��д" Then
    Text9.Text = "����"
End If
If Trim(Text10.Text) <> Trim("255.255.255.0") And Trim(Text10.Text) <> "δ��д" Then
    Text10.Text = "����"
End If
If Trim(Text11.Text) <> Trim("192.168.2.254") And Trim(Text11.Text) <> "δ��д" Then
    Text11.Text = "����"
End If
End Sub



Private Sub Command2_Click()
Text6.Text = "f0/0"
End Sub

Private Sub Command3_Click()
Text6.Text = "f1/0"
End Sub


Private Sub Command4_Click()
Text7.Text = "f1/0"
End Sub

Private Sub Command5_Click()
Text7.Text = "f0/0"
End Sub

Private Sub Command6_Click()
'Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
pc2ip.Text = ""
pc2ym.Text = ""
pc2wg.Text = ""
For i = 0 To 999
    If List1.List(i) <> "" Or List2.List(i) <> "" Then
        List1.List(i) = ""
        List2.List(i) = ""
    End If
Next i
End Sub

Private Sub Form_Load()
f1 = Frame1.Left
f2 = Frame2.Left
f3 = Frame3.Left
f4 = Frame4.Left
f5 = Frame5.Left
f6 = Frame6.Left
l1 = List1.Left
End Sub

Private Sub HScroll1_Change() 'change���

End Sub

Private Sub HScroll1_Scroll() 'scroll����
Dim a As Integer
a = HScroll1.Value
Frame1.Left = f1 + a
Frame2.Left = f2 + a
Frame3.Left = f3 + a
Frame4.Left = f4 + a
Frame5.Left = f5 + a
Frame6.Left = f6 + a
List1.Left = l1 + a
End Sub

