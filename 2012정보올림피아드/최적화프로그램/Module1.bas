Attribute VB_Name = "Module1"
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
    szExeFile As String * 260&
End Type
Private Const TH32CS_SNAPPROCESS As Long = 2&
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" _
     (ByVal lFlags As Long, lProcessID As Long) As Long
     
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" _
    (ByVal mSnapShot As Long, uProcess As PROCESSENTRY32) As Long
    
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" _
    (ByVal mSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, _
    ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private ProcessList(100, 2) As String
Private Sub KillProcessById(ByVal p_lngProcessId As Long)
  Dim lnghProcess As Long
  Dim lngReturn As Long
  lnghProcess = OpenProcess(1&, -1&, p_lngProcessId)
  lngReturn = TerminateProcess(lnghProcess, 0&)
End Sub
 
Public Sub KillProcess(ByVal ProcessName As String)
  Dim uProcess As PROCESSENTRY32
  Dim mSnapShot As Long
  Dim mName As String
  Dim i As Integer
  Dim pi As Integer
  Dim dummy As Integer
  pi = 1
  DoEvents
  uProcess.dwSize = Len(uProcess)
    
  mSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
       ' If failure -1 (false)
  If mSnapShot Then
        mresult = ProcessFirst(mSnapShot, uProcess)
            ' If failure -1 (false)
  
        Do While mresult
             i = InStr(1, uProcess.szExeFile, Chr(0))
             mName = LCase$(Left$(uProcess.szExeFile, i - 1))
             ProcessList(pi, 0) = uProcess.th32ProcessID
             ProcessList(pi, 1) = uProcess.th32ParentProcessID
             ProcessList(pi, 2) = mName
'Debug.Print mName
             mresult = ProcessNext(mSnapShot, uProcess)
             pi = pi + 1
        Loop
        
  End If
  
  For i = 1 To 100
    If ProcessList(i, 0) <> "0" Then
'       If InStr(1, ProcessList(i, 2), "iexplore.exe") = 0 Then
       If LCase(Trim(ProcessList(i, 2))) = LCase(ProcessName) Then
          KillProcessById (ProcessList(i, 0))
       End If
    End If
  Next i
End Sub



