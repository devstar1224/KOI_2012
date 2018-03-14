VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "ProcessCleanr"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4965
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4440
      Top             =   0
   End
   Begin VB.ListBox List1 
      Height          =   2580
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Const fade As Integer = 8
Dim alpha As Integer
Dim exitable As Boolean
Dim Protect_Process() As String
Private Sub Command2_Click()
Dim i As Long, j As Long, CC As Boolean, Temp() As String

If MsgBox("작업하던 내용이 종료될 수 있습니다." & vbCrLf & "괜찮으십니까?", vbInformation + vbYesNo, "[알림]") = vbYes Then
    For i = 0 To List1.ListCount - 1
        For j = 0 To 12
            If List1.List(i) = Protect_Process(j) Then
                CC = True
            End If
        Next j
        
        If CC = False Then
            KillProcess List1.List(i)
        End If
        
        CC = False
    Next i
    
    List1.Clear
    Temp = Split(CollectionProccess, "%")
    
    For i = 1 To UBound(Temp)
        List1.AddItem Temp(i)
    Next i
End If
End Sub
Private Sub Command1_Click()
Dim Temp() As String, i As Long
List1.Clear
Temp = Split(CollectionProccess, "%")
    
For i = 1 To UBound(Temp)
    List1.AddItem Temp(i)
Next i
End Sub
Private Sub Form_Load()
Dim Temp() As String, i As Long
Temp = Split(CollectionProccess, "%")

For i = 1 To UBound(Temp)
    List1.AddItem Temp(i)
Next i

ReDim Protect_Process(12)

Protect_Process(0) = "NateOnMain.exe"
Protect_Process(1) = "iexplorer.exe"
Protect_Process(2) = "explorer.exe"
Protect_Process(3) = "svchost.exe"
Protect_Process(4) = "ctfmon.exe"
Protect_Process(5) = "alg.exe"
Protect_Process(6) = "winlogon.exe"
Protect_Process(7) = "csrss.exe"
Protect_Process(8) = "smss.exe"
Protect_Process(9) = "spoolsv.exe"
Protect_Process(10) = "services.exe"
Protect_Process(11) = App.EXEName & ".exe"
Protect_Process(12) = "VB6.EXE"

Command1.Caption = "&New"
Command2.Caption = "&Clean"
End Sub

