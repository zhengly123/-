VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����/��ʾ����"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   2880
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command4 
      Caption         =   "�����ý���"
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��ʾ�ô���"
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���ظô���"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "��ϸ��Ϣ"
      Height          =   975
      Left            =   480
      TabIndex        =   3
      Top             =   960
      Width           =   3495
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   120
         TabIndex        =   5
         Text            =   "·��"
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "PID"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��þ��"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Text            =   "Ҫ���صĴ����� �� �ޱ��� - ���±�"
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "���"
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long 'SetWindowPos��������
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long 'FindWindow��������
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long 'GetWindowThreadProcessId��������
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long 'OpenProcess��������
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Const SWP_HIDEWINDOW = &H80 'SetWindowPos�����е����ش��峣��
Private Const SWP_SHOWWINDOW = &H40 'SetWindowPos�����е���ʾ���峣��
Private Const SWP_NOMOVE = &H2 'SetWindowPos�����еĲ��ƶ�����
Private Const SWP_NOSIZE = &H1 'SetWindowPos�����еĲ��ı��С����
Dim WindowHandle As Long '����WindowHandle����,������
Dim PID As Long '����PID����,����ô����������̵�PID
Dim Address As String '����address����,����ý��̵�ַ

Public Function GetProcessPathByProcessID(PID As Long) As String
On Error GoTo Z
Dim cbNeeded As Long
Dim szBuf(1 To 250) As Long
Dim Ret As Long
Dim szPathName As String
Dim nSize As Long
Dim hProcess As Long
    hProcess = OpenProcess(&H400 Or &H10, 0, PID)
    If hProcess <> 0 Then
        Ret = EnumProcessModules(hProcess, szBuf(1), 250, cbNeeded)
        If Ret <> 0 Then
            szPathName = Space(260)
            nSize = 500
            Ret = GetModuleFileNameExA(hProcess, szBuf(1), szPathName, nSize)
            GetProcessPathByProcessID = Left(szPathName, Ret)
        End If
    End If
    Ret = CloseHandle(hProcess)
    If GetProcessPathByProcessID = "" Then
        GetProcessPathByProcessID = "SYSTEM"
    End If
Exit Function
Z:
End Function

Private Sub Command1_Click()
    WindowHandle = FindWindow(vbNullString, "" & Text1.Text & "") '����������WindowHandle
    Label1.Caption = "���̾����" & WindowHandle & ""
    GetWindowThreadProcessId WindowHandle, PID  '�õ�����ID����PID
    Label2.Caption = "����PID��" & PID & ""
    Address = GetProcessPathByProcessID(PID) '�õ�·������Address
    Text2.Text = "����·����" & Address & ""
End Sub

Private Sub Command2_Click()
    SetWindowPos WindowHandle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_HIDEWINDOW '���ظþ���Ĵ���
End Sub

Private Sub Command3_Click()
    SetWindowPos WindowHandle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW  '��ʾ�þ���Ĵ���
End Sub

Private Sub Command4_Click()
    Shell "taskkill /f /pid " & PID & ""
End Sub

Private Sub Text2_DblClick()
    Shell "Explorer /select, " & Address, vbNormalFocus
End Sub
