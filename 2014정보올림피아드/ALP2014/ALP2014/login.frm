VERSION 5.00
Begin VB.Form login 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   1695
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin Project1.ctlSkin ctlSkin1 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   2990
      Caption         =   "//SCR// - ALP2014"
      Resize          =   0   'False
      Begin VB.CommandButton Command1 
         Caption         =   "로그인"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3360
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         IMEMode         =   3  '사용 못함
         Left            =   840
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "힌트:"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "PW:"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   840
         Width           =   975
      End
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If GetSetting("1", "3", "pw") = Text1.Text Then
MsgBox "로그인 완료.", vbInformation, "알림"
Unload Me
main.Show
Else
MsgBox "비밀번호가 틀립니다", vbInformation, "알림"
End If
End Sub
Private Sub Form_Load()
Label2.Caption = "힌트:" & GetSetting("1", "2", "hint")
End Sub

Private Sub Form_Unload(Cancel As Integer)
setting.Show
End Sub

