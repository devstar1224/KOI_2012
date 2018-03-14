VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Cleanr"
   ClientHeight    =   2535
   ClientLeft      =   9765
   ClientTop       =   5490
   ClientWidth     =   1995
   FillColor       =   &H8000000E&
   ForeColor       =   &H8000000E&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   1995
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command5 
      Caption         =   "Internet Reg Delete"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Help"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MemoryClean"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ProcessClean"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Internet Speed Up"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If MsgBox("레지스트리를 등록하시겟습니까?", vbQuestion + vbYesNo, "물음") = vbYes Then
Set ws = CreateObject("WScript.Shell")
ws.regWrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\AFD\Parameters\LargeBufferSize", "10240", "REG_DWORD"
ws.regWrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\AFD\Parameters\MediumBufferSize", "6016", "REG_DWORD"
ws.regWrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\AFD\Parameters\SmallBufferSize", "1024", "REG_DWORD"
ws.regWrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\AFD\Parameters\TransmitWorker", "64", "REG_DWORD"
ws.regWrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\EnablePMTUDiscovery", "1", "REG_DWORD"
ws.regWrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Tcp1323Opts", "0", "REG_DWORD"
ws.regWrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\TcpMaxDupAcks", "2", "REG_DWORD"
ws.regWrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\TcpWindowSize", "64240", "REG_DWORD"
ws.regWrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\DefaultTTL", "64", "REG_DWORD"
ws.regWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\MaxConnectionsPer1_0Serve", "64", "REG_DWORD"
ws.regWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\MaxConnectionsPerServer", "64", "REG_DWORD"
MsgBox "레지스트리를 썼습니다.", vbInformation, "알림"
Else
Close
End If
End Sub

Private Sub Command2_Click()
Form1.Show

End Sub

Private Sub Command3_Click()
moduleMemoryClean.Main

End Sub

Private Sub Command4_Click()

MsgBox "Internet Speed Up클릭시Internet 속도가 향상된다", vbInformation, "알림"
MsgBox "Process Clean클릭시 화면이 이동되면서(Form2)에서New를 눌를시 새로고침,Clean클릭시 Process가 최적화가된다.※알림 블루 스크린이 뜰수있음 그러나 컴퓨터에는 지장이 없음.", vbInformation, "알림"
MsgBox "Memory Clean클릭시 즉시 memory가 최적화가된다", vbInformation, "알림"
MsgBox "Internet Reg Delete 클릭시Internet Speed Up 을 클릭한 레지스트리(Reg)값이 삭제된다", vbInformation, "알림"
End Sub

Private Sub Command5_Click()
If MsgBox("Internet Speed Up 에서등록한 레지스트리를 삭제하시겟습니까?", vbQuestion + vbYesNo, "물음") = vbYes Then
On Error Resume Next
Set ws = CreateObject("WScript.Shell")

ws.RegDelete "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\AFD\Parameters\LargeBufferSize"

ws.RegDelete "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\AFD\Parameters\MediumBufferSize"

ws.RegDelete "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\AFD\Parameters\SmallBufferSize"


ws.RegDelete "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\AFD\Parameters\TransmitWorker"

ws.RegDelete "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\EnablePMTUDiscovery"

ws.RegDelete "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Tcp1323Opts"

ws.RegDelete "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\TcpMaxDupAcks"

ws.RegDelete "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\TcpWindowSize"

ws.RegDelete "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\DefaultTTL"

ws.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\MaxConnectionsPer1_0Serve"

ws.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\MaxConnectionsPerServer"
MsgBox "레지스트리를 제거했습니다", vbInformation, "알림"
Else
Close
End If
End Sub

