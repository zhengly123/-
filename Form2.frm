VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   Caption         =   "�����"
   ClientHeight    =   5265
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4470
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   4470
   StartUpPosition =   3  '����ȱʡ
   Begin VB.OptionButton Option7 
      Caption         =   "���汾"
      Enabled         =   0   'False
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton again 
      Caption         =   "��������"
      Height          =   375
      Left            =   3360
      TabIndex        =   17
      Top             =   4680
      Width           =   975
   End
   Begin VB.OptionButton Option6 
      Caption         =   "��ֹ����ħ��"
      Enabled         =   0   'False
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton zhuanhuan 
      Caption         =   "ת��ΪIP"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox zuohao 
      Height          =   270
      Left            =   360
      TabIndex        =   12
      Text            =   "����"
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox hua 
      Height          =   1215
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "Form2.frx":29C12
      Top             =   2520
      Width           =   1455
   End
   Begin VB.OptionButton Option999 
      Caption         =   "�����Ի�"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton close 
      Caption         =   "�Ͽ�����"
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   840
      Width           =   975
   End
   Begin VB.OptionButton Option5 
      Caption         =   "��������"
      Enabled         =   0   'False
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   3600
      Width           =   2295
   End
   Begin VB.OptionButton Option4 
      Caption         =   "�ر�ħ�޽���"
      Enabled         =   0   'False
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   3240
      Width           =   1695
   End
   Begin VB.OptionButton Option3 
      Caption         =   "��ֹ����Ϸ��ʾ"
      Enabled         =   0   'False
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2880
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      Caption         =   "�ָ�ѧ����"
      Enabled         =   0   'False
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2520
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "�رռ����"
      Enabled         =   0   'False
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox zhuangtai 
      Height          =   270
      Left            =   360
      TabIndex        =   3
      Text            =   "״̬"
      Top             =   1440
      Width           =   3135
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3600
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton fasong 
      Caption         =   "����"
      Height          =   420
      Left            =   1320
      TabIndex        =   2
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox ip 
      Height          =   270
      Left            =   360
      TabIndex        =   1
      Text            =   "IP��ַ"
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton lianjie 
      Caption         =   "����"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   840
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "����"
      Height          =   1815
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   3975
      Begin VB.Label Label8 
         Caption         =   "�汾��V1.0"
         Height          =   255
         Left            =   3000
         TabIndex        =   15
         Top             =   0
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private c As Integer                                            '������
Private cishu As Integer                                        '������
Private d As Integer                                            '������

Private Sub again_Click()
Unload Me
Me.Show
End Sub

Private Sub hua_Click()
cishu = cishu + 1                                            'ɾ������
If cishu = 1 Then hua.Text = ""
End Sub

Private Sub ip_click()
c = c + 1                                                    'ɾ������
If c = 1 Then ip = ""
End Sub

Private Sub lianjie_Click()
On Error GoTo F
    Winsock1.RemoteHost = ip.Text                                '��Ȼ�Ǿ��������� �ǾͲ���Ҫ�˿�ӳ�� �ҵ����ھ������ڵ�ip���У�����Ӧ���Ƕ�̬���ġ�����������text��
    Winsock1.RemotePort = 2012                                   '����Ҫ���ӵ�2012�˿�  (Ҫ�Ϳͻ������õ�һ��)
    Winsock1.Connect                                             '����
    zhuangtai.Text = "������..."
Exit Sub
F:
    zhuangtai.Text = "������"
End Sub

Private Sub fasong_Click()
    ans = MsgBox("��ȷ��Ҫִ�в�����", vbYesNo)                  '����Ի���
    If ans = vbYes Then                                          '�û����¡��ǡ�
        If Option1.Value = True Then Winsock1.SendData "a"           '��ͻ��˷������� "a"
        If Option2.Value = True Then Winsock1.SendData "b"           '��ͻ��˷������� "b"
        If Option3.Value = True Then Winsock1.SendData "c"           '��ͻ��˷������� "c"
        If Option4.Value = True Then Winsock1.SendData "d"           '��ͻ��˷������� "d"
        If Option5.Value = True Then Winsock1.SendData "s"           '��ͻ��˷������� "s"
        If Option6.Value = True Then Winsock1.SendData "e"           '��ͻ��˷������� "e"
        If Option7.Value = True Then Winsock1.SendData "banb"        '��ͻ��˷������� "banb"
        If Option999.Value = True Then Winsock1.SendData "word" & hua & "" '��ͻ��˷�������Ҫ˵�Ļ�
    End If
End Sub
 
Private Sub exit_Click()                                     '�Ͽ�����
    Winsock1.close
    zhuangtai.Text = "�ѶϿ�"
End Sub

Private Sub Winsock1_Close()
    Option1.Enabled = False
    Option2.Enabled = False
    Option3.Enabled = False
    Option4.Enabled = False
    Option5.Enabled = False
    Option6.Enabled = False
    Option7.Enabled = False
    Option999.Enabled = False
zhuangtai.Text = "�ѶϿ�"
End Sub

Private Sub Winsock1_Connect()                               '��������
    Option1.Enabled = True                                       '��ѡȫ������
    Option2.Enabled = True
    Option3.Enabled = True
    Option4.Enabled = True
    Option5.Enabled = True
    Option6.Enabled = True
    Option7.Enabled = True
    Option999.Enabled = True
    zhuangtai.Text = "������" & ip & ""
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim strget As String
    Winsock1.GetData strget                                          '�ѽ��ܵ������ݷŵ�strget��
    Select Case Left(strget, 4)
    Case "cuow"
        MsgBox "�ͻ��˽Ӵ������ݳ���"
    Case "banb"
        MsgBox "�ÿͻ��˵İ汾Ϊ" & Mid(strget, 5) & ""
    Case "debu"
        MsgBox "ͨѶ����"
    Case Else
        MsgBox "�������ݳ���"
    End Select
End Sub

Private Sub zhuanhuan_Click()
On Error Resume Next
    c = c + 1
    aaa = Mid(zuohao.Text, 1, 1)
    bbb = Mid(zuohao.Text, 2, 2)
    If Asc(aaa) <= 70 Then ccc = ((Asc(aaa) - 64) * 10 - 1) + Val(bbb)
    If Asc(aaa) > 70 Then ccc = ((Asc(aaa) - 96) * 10 - 1) + Val(bbb)
    ip.Text = "192.168.4." & ccc & ""
End Sub

Private Sub zuohao_Click()
    d = d + 1                                                      'ɾ������
    If d = 1 Then
        zuohao = ""
        zhuanhuan.Enabled = True
    End If
End Sub
