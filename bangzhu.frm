VERSION 5.00
Begin VB.Form bangzhu 
   Caption         =   "帮助"
   ClientHeight    =   2535
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "bangzhu.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   2535
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "本机信息"
      Height          =   1095
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   3735
      Begin VB.Label Label2 
         Caption         =   "无法获取本机信息"
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Label Label4 
      Caption         =   "by xY"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "甲虫版3.0正式版（C）                 F11和F12以及F10，仅仅3个键，还等什么，快"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "bangzhu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim aa As String
Dim strLocalIP As String
Dim winIP As Object
aa = aa & "本机电脑名称:" & Environ("computername") & vbCrLf
aa = aa & "本机用户名称:" & Environ("username") & vbCrLf
Set winIP = CreateObject("MSWinsock.Winsock")
strLocalIP = winIP.LocalIP
Label2.Caption = aa & "本机IP:" & strLocalIP
Label3.Caption = "版本号" & App.Major & "." & App.Minor & "." & App.Revision & ""
End Sub
