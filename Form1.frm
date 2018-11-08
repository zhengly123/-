VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "甲虫版3.0"
   ClientHeight    =   5175
   ClientLeft      =   1860
   ClientTop       =   1485
   ClientWidth     =   4845
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   4845
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   480
   End
   Begin VB.CommandButton over 
      Caption         =   "自我销毁"
      Height          =   660
      Left            =   0
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton queding 
      Caption         =   "确定"
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   600
      Width           =   615
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton gb 
      Caption         =   "F6 关闭指定进程"
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox jincheng 
      Height          =   270
      Left            =   960
      TabIndex        =   12
      Text            =   "War3.exe"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton zy 
      Caption         =   "F5 显示桌面"
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton xz 
      Caption         =   " F10 卸载360  (必须的)"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   0
   End
   Begin VB.CommandButton C2 
      Caption         =   "F12 恢复"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "托盘化（推荐）"
      Height          =   495
      Left            =   2520
      TabIndex        =   6
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   3960
      TabIndex        =   5
      Text            =   "reg delete ""HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\NetClassStu2007.exe"" /f"
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   3120
      TabIndex        =   4
      Text            =   $"Form1.frx":29C12
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox guanliyuan 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   840
      PasswordChar    =   "*"
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "高级密码"
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton C1 
      Caption         =   "F11 关闭"
      Height          =   495
      Left            =   840
      TabIndex        =   9
      ToolTipText     =   "关掉360的一切哦 不然没用滴"
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "主功能"
      Height          =   1575
      Left            =   720
      TabIndex        =   11
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Frame Frame2 
      Caption         =   "New!"
      ForeColor       =   &H000000FF&
      Height          =   2055
      Left            =   720
      TabIndex        =   14
      Top             =   1320
      Width           =   3375
      Begin VB.CommandButton moshou 
         Caption         =   "Warcraft III"
         Height          =   255
         Left            =   1320
         TabIndex        =   21
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton showit 
         Caption         =   "F3 显示窗体"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1920
         TabIndex        =   19
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton hideit 
         Caption         =   "F2 隐藏窗体"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox chuangti 
         Height          =   270
         Left            =   240
         TabIndex        =   16
         Text            =   "要隐藏的窗体名 例 无标题 - 记事本"
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label Label4 
         Caption         =   "常用窗体名："
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label3 
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   3360
         Y1              =   840
         Y2              =   840
      End
   End
   Begin VB.Label Label2 
      Caption         =   "    请先阅读帮助，Esc为推出。建议最小化(托盘中和qq一样)，不要关闭"
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000010&
      Caption         =   "    新功能，使用方法很简单。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.Menu afae 
      Caption         =   "设置"
   End
   Begin VB.Menu qwerqwer 
      Caption         =   "工具"
      Begin VB.Menu afawef 
         Caption         =   "隐藏/显示窗体"
      End
   End
   Begin VB.Menu qwreqwerq 
      Caption         =   "帮助"
      Begin VB.Menu help 
         Caption         =   "关于甲虫版（&A）"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'发行修改部分
'1、自我销毁按钮改为不可见
'2、删除退出按钮
'3、
'------热键部分
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Type NOTIFYICONDATA
cbSize As Long
hwnd As Long
uId As Long
uFlags As Long
uCallBackMessage As Long
hIcon As Long
szTip As String * 64
End Type
Dim tray As NOTIFYICONDATA
'------隐藏显示窗体部分
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long 'SetWindowPos函数声明
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long 'FindWindow函数声明
Private Const SWP_HIDEWINDOW = &H80 'SetWindowPos函数中的隐藏窗体常数
Private Const SWP_SHOWWINDOW = &H40 'SetWindowPos函数中的显示窗体常数
Private Const SWP_NOMOVE = &H2 'SetWindowPos函数中的不移动常数
Private Const SWP_NOSIZE = &H1 'SetWindowPos函数中的不改变大小常数
Dim WindowHandle As Long '声明WindowHandle变量,储存句柄
'------托盘化
Const NIM_ADD = &H0
Const NIM_DELETE = &H2
Const NIF_ICON = &H2
Const NIF_MESSAGE = &H1
Const NIF_TIP = &H4
Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDBLCLK = &H203
'------模拟鼠标键盘
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long) '键盘模拟
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long) '鼠标模拟
'------其他
Dim pjishuqi As Long                                                                                 '管理员密码计数器

Private Sub afawef_Click()
    Form3.Show
End Sub

Private Sub chuangti_Change()
    WindowHandle = FindWindow(vbNullString, "" & chuangti.Text & "") '窗体句柄存入WindowHandle
    If WindowHandle > 0 Then
        hideit.Enabled = True
        showit.Enabled = True
    Else
        hideit.Enabled = False
        showit.Enabled = False
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)                               '按关闭按钮时隐藏
    If App.PrevInstance = True Then
        End
    Else
        Me.Hide
        Cancel = vbCancel
        Timer1.Enabled = False
        Timer2.Enabled = True
    End If
End Sub

Private Sub C1_Click()
    Shell "taskkill /f /t /im NetClassStu2007.exe", vbHide '结束进程
    Shell "taskkill /f /t /im DF5Serv.exe", vbHide '结束冰点还原
    Shell "reg add ""HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\NetClassStu2007.exe"" /v debugger /t reg_sz /d debugfile.exe /f", vbHide '禁止启动
    C2.Enabled = True
    C1.Enabled = False
End Sub

Private Sub C2_Click()
    Shell "reg delete ""HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\NetClassStu2007.exe"" /f", vbHide '恢复运行代码
    C2.Enabled = False
    C1.Enabled = True
End Sub

Private Sub C3_Click()
End '测试时使用
'发行时转为正码If App.PrevInstance = False Then
'Form1.Visible = False '退出时后台运行
'Else
'End
'End If
End Sub

Private Sub gb_Click()
    If jincheng.Text = App.EXEName Then
        MsgBox "慎重操作，你再试一次看"
    Else
        Shell "cmd /c taskkill /f /t /im " & jincheng & "", vbHide       '关闭指定程序
    End If
End Sub

Private Sub dueding_Click()
    If Text3.Text = "123123" Then                                    '使用密码设定
        C1.Enabled = True
    Else
        MsgBox "密码错误，请重新输入"
        Text3.Text = ""
    End If
End Sub

Private Sub Form_Load()
    tray.cbSize = Len(tray)                                             '热键设置
    tray.uId = vbNull
    tray.hwnd = Me.hwnd
    tray.uFlags = NIF_TIP Or NIF_MESSAGE Or NIF_ICON
    tray.uCallBackMessage = WM_MOUSEMOVE
    tray.hIcon = Me.Icon
    tray.szTip = vbNullChar
    Shell_NotifyIcon NIM_ADD, tray
    Me.Hide
    On Error Resume Next
        Winsock1.Bind 2012                                              '测试2012端口
        If Err = 0 Then                                                 '端口未被占用
        Winsock1.LocalPort = 2012                                       '本地打开的网络端口 客户端要连接到这个端口
        Winsock1.Listen                                                 '保持监听状态
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim msg As Long
    msg = x / 15
    If msg = WM_LBUTTONDBLCLK Then
        Me.Show
        Shell_NotifyIcon NIM_DELETE, tray
    End If
End Sub

Private Sub Command1_Click()                                '托盘化,看不懂
    tray.cbSize = Len(tray)
    tray.uId = vbNull
    tray.hwnd = Me.hwnd
    tray.uFlags = NIF_TIP Or NIF_MESSAGE Or NIF_ICON
    tray.uCallBackMessage = WM_MOUSEMOVE
    tray.hIcon = Me.Icon
    tray.szTip = vbNullChar
    Shell_NotifyIcon NIM_ADD, tray
    Me.Hide
End Sub


Private Sub guanliyuan_KeyPress(KeyAscii As Integer)
    If KeyAscii > Asc("9") Then                                    '禁止输入英文
        KeyAscii = 0
    End If
End Sub

Private Sub help_Click()
bangzhu.Show
End Sub

Private Sub hideit_Click()
    SetWindowPos WindowHandle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_HIDEWINDOW '隐藏该句柄的窗体
End Sub

Private Sub moshou_Click()
    chuangti.Text = "Warcraft III"
End Sub

Private Sub queding_Click()
    If guanliyuan.Text = "441021" Then                                 '管理员密码设定
        Text1.Visible = True: Text2.Visible = True:                        '显示指令
        Winsock1.close                                                     'winsock关闭
        Form2.Show
        guanliyuan.Text = ""
    Else
        pjishuqi = pjishuqi + 1                                            '管理员密码计数器
        guanliyuan.Text = ""
        If pjishuqi = 1 Then
            MsgBox "密码错误，请重新输入，你还有2次机会", vbOKOnly, "甲虫版3.0" '提示次数
        End If
        If pjishuqi = 2 Then
            MsgBox "密码错误，请重新输入，你还有1次机会"
        End If
        If pjishuqi = 3 Then
            MsgBox "密码错误3次，退出"                                         '错误3次关机
            Shell "shutdown -s -f -t 30"
            End
        End If
    End If
End Sub

Private Sub showit_Click()
    SetWindowPos WindowHandle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW  '显示该句柄的窗体
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next                                             '热键
    If GetAsyncKeyState(vbKeyNumlock) Then Me.Show
    If GetAsyncKeyState(vbKeyF2) Then hideit_Click
    If GetAsyncKeyState(vbKeyF3) Then showit_Click
    If GetAsyncKeyState(vbKeyF5) Then zy_Click
    If GetAsyncKeyState(vbKeyF6) Then gb_Click
    If GetAsyncKeyState(vbKeyF10) Then xz_Click
    If GetAsyncKeyState(vbKeyF11) Then C1_Click
    If GetAsyncKeyState(vbKeyF12) Then C2_Click
    If GetAsyncKeyState(vbKeyEnd) Then End
End Sub

Private Sub over_Click()
    On Error Resume Next
        s = CurDir '当前目录
        '保证目录最后的字符为 "\"
        If Right(s, 1) <> "\" Then s = s & "\"
    '在当前目录下创建bat文件
    Open s & "kill.bat" For Output As #1
        Print #1, ":redel"
        Print #1, "del " & Chr(34) & s & App.EXEName & ".exe" & Chr(34)
        Print #1, "if exist " & Chr(34) & s & App.EXEName & ".exe" & Chr(34) & " goto redel"
        Print #1, "del %0"
        Print #1,
        Close #1
    Shell Chr(34) & s & "kill.bat" & Chr(34), vbHide
    End
End Sub

Private Sub Timer2_Timer()
    On Error Resume Next                                             '热键
    If GetAsyncKeyState(vbKeyNumlock) Then
        Me.Show
        Timer2.Enabled = False
        Timer1.Enabled = True
    End If
End Sub

Private Sub xz_Click()
    Shell "C:\Program Files\360\360safe\uninst.exe"                  '调用360卸载
End Sub

Private Sub zy_Click()
    CreateObject("Shell.Application").ToggleDesktop                  '显示桌面
End Sub
Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)  '受到连接请求
    Winsock1.close
    Winsock1.Accept requestID                                        '接受请求
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)       '收到数据时
    Dim strget As String
    Winsock1.GetData strget                                          '把接受到的数据放到strget中
        Select Case Left(strget, 4)                              '截取前四个字符
        Case "word"
            strget = Mid(strget, 5)
            MsgBox "" & strget & "", vbOKOnly, "甲虫版3.0"
        Case "banb"
            Winsock1.SendData "banb" & App.Major & "." & App.Minor & "." & App.Revision & ""
        Case "debu"
            Winsock1.SendData "debu"
        Case Else
            Select Case strget
                Case "a"
                    Shell "shutdown -s -t 0"                    '接受数据为a，则关机
                Case "b"
                    C2_Click                                    '如果数据为b，则C2_click
                Case "c"
                    MsgBox "今天别玩游戏，不关掉游戏后果自负", vbOKOnly, "甲虫版3.0"    '接收数据为c，则弹出提示
                Case "d"
                    Shell "taskkill /f /im War3.exe"            '接收数据为d，则关闭魔兽进程
                Case "e"
                    Shell "reg add ""HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\War3.exe"" /v debugger /t reg_sz /d debugfile.exe /f", vbHide '接收数据为e，则禁止启动魔兽
                Case "s"
                    over_Click                                  '接收数据为s，则自我销毁
                Case Else
                    Winsock1.SendData "cuow"
            End Select
        End Select
End Sub

