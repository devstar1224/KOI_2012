Attribute VB_Name = "moduleMemoryClean"
' 다시 가입한 러브모리(lovemomory)

Private Const MAX_PATH As Long = 260

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Private Declare Function CreateToolhelp32Snapshot Lib "kernel32.dll" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32.dll" (ByVal hSnapshot As Long, ByRef lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32.dll" (ByVal hSnapshot As Long, ByRef lppe As PROCESSENTRY32) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

Private Declare Function EmptyWorkingSet Lib "psapi.dll" (ByVal hProcess As Long) As Long

Private Const TH32CS_SNAPPROCESS As Long = &H2

Private Const PROCESS_QUERY_INFORMATION As Long = (&H400)
Private Const PROCESS_SET_QUOTA As Long = (&H100)

Sub Main()
    Dim pe32 As PROCESSENTRY32, snap As Long
    pe32.dwSize = Len(pe32)
    snap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    Dim r As Long
    r = Process32First(snap, pe32)
    Dim pid As Long, hProcess As Long, cnt As Long
    Do While r
        pid = pe32.th32ProcessID
        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_SET_QUOTA, 0&, pe32.th32ProcessID)
        If Not CleanMem(hProcess) = 0 Then cnt = cnt + 1
        r = Process32Next(snap, pe32)
    Loop
    CloseHandle snap
    MsgBox cnt & "개 프로세스의 메모리 정리 됨"
End Sub

Private Function CleanMem(hProcess As Long) As Long
    CleanMem = EmptyWorkingSet(hProcess)
    CloseHandle hProcess
End Function
