VERSION 5.00
Begin VB.Form bangzhu 
   Caption         =   "����"
   ClientHeight    =   2535
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "bangzhu.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   2535
   ScaleWidth      =   4560
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame1 
      Caption         =   "������Ϣ"
      Height          =   1095
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   3735
      Begin VB.Label Label2 
         Caption         =   "�޷���ȡ������Ϣ"
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
      Caption         =   "�׳��3.0��ʽ�棨C��                 F11��F12�Լ�F10������3����������ʲô����"
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
aa = aa & "������������:" & Environ("computername") & vbCrLf
aa = aa & "�����û�����:" & Environ("username") & vbCrLf
Set winIP = CreateObject("MSWinsock.Winsock")
strLocalIP = winIP.LocalIP
Label2.Caption = aa & "����IP:" & strLocalIP
Label3.Caption = "�汾��" & App.Major & "." & App.Minor & "." & App.Revision & ""
End Sub
