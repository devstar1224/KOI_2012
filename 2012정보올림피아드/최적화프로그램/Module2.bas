Attribute VB_Name = "Module2"
Option Explicit
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32.dll" ( _
    ByVal dwFlags As Long, _
    ByVal th32ProcessID As Long _
) As Long
Public Declare Function Process32First Lib "kernel32.dll" ( _
    ByVal hSnapShot As Long, _
    ByRef lppe As PROCESSENTRY32 _
) As Long
Public Declare Function Process32Next Lib "kernel32.dll" ( _
    ByVal hSnapShot As Long, _
    ByRef lppe As PROCESSENTRY32 _
) As Long
Public Declare Function CloseHandle Lib "kernel32.dll" ( _
    ByVal hObject As Long _
) As Long
Public Const MAX_PATH& = 260&
Public Const INVALID_HANDLE_VALUE& = -1&
Public Const TH32CS_SNAPPROCESS& = 2&
Public Type PROCESSENTRY32
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
Public Function CollectionProccess() As String
Dim Temp As String
Dim pe As PROCESSENTRY32, hSnapShot As Long, lRet As Long, ImagePath As String
    
    pe.dwSize = Len(pe)
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    
    If hSnapShot <> INVALID_HANDLE_VALUE Then
        lRet = Process32First(hSnapShot, pe)
        
        Do While lRet
            ImagePath = Split(pe.szExeFile, vbNullChar, 2)(0)
            ImagePath = Mid$(ImagePath, InStrRev(ImagePath, "\") + 1)
            Temp = Temp & ImagePath & "%"
            lRet = Process32Next(hSnapShot, pe)
        Loop
        CloseHandle hSnapShot
        Temp = Mid(Temp, 1, Len(Temp) - 1)
    End If
    
    CollectionProccess = Temp
    'Sock.SendData "%Process%" & Temp
End Function





