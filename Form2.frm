VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   Caption         =   "服务端"
   ClientHeight    =   5265
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4470
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   4470
   StartUpPosition =   3  '窗口缺省
   Begin VB.OptionButton Option7 
      Caption         =   "检查版本"
      Enabled         =   0   'False
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton again 
      Caption         =   "重新连接"
      Height          =   375
      Left            =   3360
      TabIndex        =   17
      Top             =   4680
      Width           =   975
   End
   Begin VB.OptionButton Option6 
      Caption         =   "禁止开启魔兽"
      Enabled         =   0   'False
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton zhuanhuan 
      Caption         =   "转换为IP"
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
      Text            =   "座号"
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
      Caption         =   "单方对话"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton close 
      Caption         =   "断开连接"
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   840
      Width           =   975
   End
   Begin VB.OptionButton Option5 
      Caption         =   "自我销毁"
      Enabled         =   0   'False
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   3600
      Width           =   2295
   End
   Begin VB.OptionButton Option4 
      Caption         =   "关闭魔兽进程"
      Enabled         =   0   'False
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   3240
      Width           =   1695
   End
   Begin VB.OptionButton Option3 
      Caption         =   "禁止玩游戏提示"
      Enabled         =   0   'False
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2880
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      Caption         =   "恢复学生端"
      Enabled         =   0   'False
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2520
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "关闭计算机"
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
      Text            =   "状态"
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
      Caption         =   "发送"
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
      Text            =   "IP地址"
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton lianjie 
      Caption         =   "连接"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   840
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "连接"
      Height          =   1815
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   3975
      Begin VB.Label Label8 
         Caption         =   "版本号V1.0"
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
Private c As Integer                                            '计数器
Private cishu As Integer                                        '计数器
Private d As Integer                                            '计数器

Private Sub again_Click()
Unload Me
Me.Show
End Sub

Private Sub hua_Click()
cishu = cishu + 1                                            '删除文中
If cishu = 1 Then hua.Text = ""
End Sub

Private Sub ip_click()
c = c + 1                                                    '删除文中
If c = 1 Then ip = ""
End Sub

Private Sub lianjie_Click()
On Error GoTo F
    Winsock1.RemoteHost = ip.Text                                '既然是局域网互关 那就不需要端口映射 找到他在局域网内的ip就行，不过应该是动态会变的。。所以做个text好
    Winsock1.RemotePort = 2012                                   '设置要连接到2012端口  (要和客户端设置的一样)
    Winsock1.Connect                                             '连接
    zhuangtai.Text = "连接中..."
Exit Sub
F:
    zhuangtai.Text = "出错啦"
End Sub

Private Sub fasong_Click()
    ans = MsgBox("您确定要执行操作吗", vbYesNo)                  '警告对话框
    If ans = vbYes Then                                          '用户按下“是”
        If Option1.Value = True Then Winsock1.SendData "a"           '向客户端发送数据 "a"
        If Option2.Value = True Then Winsock1.SendData "b"           '向客户端发送数据 "b"
        If Option3.Value = True Then Winsock1.SendData "c"           '向客户端发送数据 "c"
        If Option4.Value = True Then Winsock1.SendData "d"           '向客户端发送数据 "d"
        If Option5.Value = True Then Winsock1.SendData "s"           '向客户端发送数据 "s"
        If Option6.Value = True Then Winsock1.SendData "e"           '向客户端发送数据 "e"
        If Option7.Value = True Then Winsock1.SendData "banb"        '向客户端发送数据 "banb"
        If Option999.Value = True Then Winsock1.SendData "word" & hua & "" '向客户端发送数据要说的话
    End If
End Sub
 
Private Sub exit_Click()                                     '断开连接
    Winsock1.close
    zhuangtai.Text = "已断开"
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
zhuangtai.Text = "已断开"
End Sub

Private Sub Winsock1_Connect()                               '连接上了
    Option1.Enabled = True                                       '单选全部开启
    Option2.Enabled = True
    Option3.Enabled = True
    Option4.Enabled = True
    Option5.Enabled = True
    Option6.Enabled = True
    Option7.Enabled = True
    Option999.Enabled = True
    zhuangtai.Text = "已连接" & ip & ""
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim strget As String
    Winsock1.GetData strget                                          '把接受到的数据放到strget中
    Select Case Left(strget, 4)
    Case "cuow"
        MsgBox "客户端接处理数据出错"
    Case "banb"
        MsgBox "该客户端的版本为" & Mid(strget, 5) & ""
    Case "debu"
        MsgBox "通讯正常"
    Case Else
        MsgBox "接收数据出错"
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
    d = d + 1                                                      '删除文中
    If d = 1 Then
        zuohao = ""
        zhuanhuan.Enabled = True
    End If
End Sub
