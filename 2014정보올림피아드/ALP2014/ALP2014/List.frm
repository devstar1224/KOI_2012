VERSION 5.00
Begin VB.Form List 
   BorderStyle     =   0  '¾øÀ½
   Caption         =   "//IDList// Automatic Login Program - ALP2014"
   ClientHeight    =   5100
   ClientLeft      =   2370
   ClientTop       =   1575
   ClientWidth     =   5460
   Icon            =   "List.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "´Ý±â"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   25
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   3000
      Top             =   3720
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "ÀÌÈ­¸éÀ» ´Ù½Ã¶Ù¿ï°æ¿ì »õ·Î°íÄ§ÀÌ µË´Ï´Ù."
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   24
      Top             =   4440
      Width           =   4335
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "5."
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   240
      TabIndex        =   23
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "12."
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   120
      TabIndex        =   22
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "11."
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   120
      TabIndex        =   21
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "10."
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   20
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "9."
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   240
      TabIndex        =   19
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "8."
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   240
      TabIndex        =   18
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "7."
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   17
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "6."
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   16
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "4."
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   15
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "3."
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   14
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "2."
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "1."
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Null"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   4080
      Width           =   5175
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Null"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   3720
      Width           =   5175
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Null"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   3360
      Width           =   5175
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Null"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   3000
      Width           =   5175
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Null"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   2640
      Width           =   5175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Null"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   2280
      Width           =   5175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Null"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1920
      Width           =   5175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Null"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   5175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Null"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1200
      Width           =   5175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Null"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Null"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Null"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const ws_ex_layered As Long = &H80000
Private Const GWL_EXSTYLE As Long = -20
Private Const LWA_ALPHA As Long = &H2
Option Explicit
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Private Declare Function GetKeyNameText Lib "user32" Alias "GetKeyNameTextA" (ByVal lParam As Long, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
  SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or ws_ex_layered
  SetLayeredWindowAttributes Me.hwnd, 0, 200, LWA_ALPHA
Timer1.Interval = setting.Text16.Text

If main.Text1.Text = "" Then
Else
Label1.Caption = "ID:" & main.Text1.Text & " / Site:" & main.Text32.Text & " /Cord:" & setting.Text1
End If

If main.Text3.Text = "" Then
Else
Label2.Caption = "ID:" & main.Text3.Text & " / Site:" & main.Text33.Text & " /Cord:" & setting.Text2
End If


If main.Text5.Text = "" Then
Else
Label3.Caption = "ID:" & main.Text5.Text & " / Site:" & main.Text31.Text & " /Cord:" & setting.Text3
End If

If main.Text7.Text = "" Then
Else
Label4.Caption = "ID:" & main.Text7.Text & " / Site:" & main.Text30.Text & " /Cord:" & setting.Text4
End If

If main.Text17.Text = "" Then
Else
Label5.Caption = "ID:" & main.Text17.Text & "/ Site:" & main.Text34.Text & " /Cord:" & setting.Text11
End If

If main.Text15.Text = "" Then
Else
Label6.Caption = "ID:" & main.Text15.Text & "/ Site:" & main.Text35.Text & " /Cord:" & setting.Text10
End If

If main.Text13.Text = "" Then
Else
Label7.Caption = "ID:" & main.Text13.Text & "/ Site:" & main.Text36.Text & " /Cord:" & setting.Text9
End If

If main.Text11.Text = "" Then
Else
Label8.Caption = "ID:" & main.Text11.Text & "/ Site:" & main.Text37.Text & " /Cord:" & setting.Text8
End If

If main.Text18.Text = "" Then
Else
Label9.Caption = "ID:" & main.Text18.Text & "/ Site:" & main.Text26.Text & " /Cord:" & setting.Text12

End If

If main.Text20.Text = "" Then
Else
Label10.Caption = "ID:" & main.Text20.Text & "/ Site:" & main.Text27.Text & " /Cord:" & setting.Text13
End If
0
If main.Text22.Text = "" Then
Else
Label11.Caption = "ID:" & main.Text22.Text & "/ Site:" & main.Text28.Text & " /Cord:" & setting.Text14
End If

If main.Text24.Text = "" Then
Else
Label12.Caption = "ID:" & main.Text24.Text & "/ Site:" & main.Text29.Text & " /Cord:" & setting.Text15
End If
Timer1.Enabled = True
End Sub

Private Sub Label10_Click()

End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub
