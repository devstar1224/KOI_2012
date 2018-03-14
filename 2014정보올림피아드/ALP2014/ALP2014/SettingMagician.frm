VERSION 5.00
Begin VB.Form SettingMagician 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   6855
   ClientLeft      =   315
   ClientTop       =   2085
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin Project1.ctlSkin ctlSkin1 
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   12091
      Caption         =   "//S.T.M//Automatic Login Program - ALP2014"
      Resize          =   0   'False
      Begin VB.Frame Frame13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Setting"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4815
         Left            =   3720
         TabIndex        =   27
         Top             =   1320
         Width           =   3375
         Begin VB.CommandButton Command14 
            Caption         =   "Setting1 저장"
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   480
            TabIndex        =   38
            Top             =   4080
            Width           =   2295
         End
         Begin VB.TextBox Text4 
            Height          =   270
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   3120
            Width           =   615
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   240
            TabIndex        =   33
            Text            =   "ex)네이버,naver.com"
            Top             =   2160
            Width           =   2895
         End
         Begin VB.TextBox Text2 
            Height          =   270
            IMEMode         =   3  '사용 못함
            Left            =   840
            PasswordChar    =   "*"
            TabIndex        =   31
            Top             =   1320
            Width           =   1935
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Left            =   840
            TabIndex        =   29
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label9 
            BackStyle       =   0  '투명
            Caption         =   "텍스트박스에 클릭을하신후 하고싶은 키를 눌러주세요"
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
            TabIndex        =   41
            Top             =   3600
            Width           =   2775
         End
         Begin VB.Label Label6 
            BackStyle       =   0  '투명
            Caption         =   "아이디 설정"
            BeginProperty Font 
               Name            =   "a피오피네모"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   37
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '투명
            Caption         =   "End +"
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
            Left            =   960
            TabIndex        =   36
            Top             =   3120
            Width           =   1335
         End
         Begin VB.Label Label4 
            BackStyle       =   0  '투명
            Caption         =   "키설정"
            BeginProperty Font 
               Name            =   "a피오피네모"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1200
            TabIndex        =   34
            Top             =   2760
            Width           =   975
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '투명
            Caption         =   "Site"
            BeginProperty Font 
               Name            =   "a피오피네모"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1320
            TabIndex        =   32
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Label2 
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
            TabIndex        =   30
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "ID:"
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   480
            TabIndex        =   28
            Top             =   840
            Width           =   1335
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Setting12"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   7200
         TabIndex        =   25
         Top             =   5760
         Width           =   3495
         Begin VB.CommandButton Command13 
            Caption         =   "Setting12 종합설정 하기"
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Setting11"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   7200
         TabIndex        =   23
         Top             =   4800
         Width           =   3495
         Begin VB.CommandButton Command12 
            Caption         =   "Setting11 종합설정 하기"
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   24
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Setting10"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   7200
         TabIndex        =   21
         Top             =   3840
         Width           =   3495
         Begin VB.CommandButton Command11 
            Caption         =   "Setting10 종합설정 하기"
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Setting9"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   7200
         TabIndex        =   19
         Top             =   2880
         Width           =   3495
         Begin VB.CommandButton Command10 
            Caption         =   "Setting9 종합설정 하기"
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   20
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Setting8"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   7200
         TabIndex        =   17
         Top             =   1920
         Width           =   3495
         Begin VB.CommandButton Command9 
            Caption         =   "Setting8 종합설정 하기"
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Setting7"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   7200
         TabIndex        =   15
         Top             =   960
         Width           =   3495
         Begin VB.CommandButton Command8 
            Caption         =   "Setting7 종합설정 하기"
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Setting6"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   13
         Top             =   5760
         Width           =   3495
         Begin VB.CommandButton Command7 
            Caption         =   "Setting6 종합설정 하기"
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Setting5"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   11
         Top             =   4800
         Width           =   3495
         Begin VB.CommandButton Command6 
            Caption         =   "Setting5 종합설정 하기"
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Setting4"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   9
         Top             =   3840
         Width           =   3495
         Begin VB.CommandButton Command5 
            Caption         =   "Setting4 종합설정 하기"
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Setting3"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   3495
         Begin VB.CommandButton Command4 
            Caption         =   "Setting3 종합설정 하기"
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Setting2"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   3495
         Begin VB.CommandButton Command3 
            Caption         =   "Setting2 종합설정 하기"
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Setting1"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   3495
         Begin VB.CommandButton Command2 
            Caption         =   "Setting1 종합설정 하기"
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   2
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.Label Label8 
         BackStyle       =   0  '투명
         Caption         =   "※ 저장을 하면 적었던 값이 사라지는게 정상입니다."
         Height          =   375
         Left            =   3720
         TabIndex        =   40
         Top             =   6240
         Width           =   3375
      End
      Begin VB.Line Line1 
         X1              =   5400
         X2              =   5400
         Y1              =   480
         Y2              =   840
      End
      Begin VB.Label Label7 
         BackStyle       =   0  '투명
         Caption         =   "※ 자신이 미리 세팅을 하였어도 이창에는 이전에있던 아이디기 표기가 되지않습니다."
         Height          =   375
         Left            =   6000
         TabIndex        =   39
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label38 
         BackStyle       =   0  '투명
         Caption         =   "※ 종합설정시 자동으로 사용체크가 됩니다."
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   720
         Width           =   4815
      End
      Begin VB.Label Label37 
         BackStyle       =   0  '투명
         Caption         =   "※ 종합설정을 할경우 기존에 설정된 세팅이 지워집니다."
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   480
         Width           =   5535
      End
   End
End
Attribute VB_Name = "SettingMagician"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Keycord.Show
End Sub


Private Sub Command10_Click()
Command14.Caption = "Setting9 저장"
End Sub

Private Sub Command11_Click()
Command14.Caption = "Setting10 저장"
End Sub

Private Sub Command12_Click()
Command14.Caption = "Setting11 저장"
End Sub

Private Sub Command13_Click()
Command14.Caption = "Setting12 저장"
End Sub

Private Sub Command14_Click()
If Command14.Caption = "Setting1 저장" Then
main.Text1.Text = Text1.Text
main.Text2.Text = Text2.Text
main.Text32.Text = Text3.Text
setting.Text1.Text = Text4.Text
SaveSetting "1", "id", "3", Text1.Text
SaveSetting "1", "pw", "3", Text2.Text
SaveSetting "1", "site", "3", Text3.Text
SaveSetting "1", "key", "3", Text4.Text
SaveSetting "ch", "1", "3", "che1"
main.Check1.value = 1
main.Hide
MsgBox "Setting1 에관한 정보를 저장하였습니다", vbInformation, "알림"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End If

If Command14.Caption = "Setting2 저장" Then
main.Text3.Text = Text1.Text
main.Text4.Text = Text2.Text
main.Text33.Text = Text3.Text
setting.Text2.Text = Text4.Text
SaveSetting "2", "id", "3", Text1.Text
SaveSetting "2", "pw", "3", Text2.Text
SaveSetting "2", "site", "3", Text3.Text
SaveSetting "2", "key", "3", Text4.Text
SaveSetting "ch", "2", "3", "che2"
main.Check2.value = 1
main.Hide
MsgBox "Setting2 에관한 정보를 저장하였습니다", vbInformation, "알림"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End If

If Command14.Caption = "Setting3 저장" Then
main.Text5.Text = Text1.Text
main.Text6.Text = Text2.Text
main.Text31.Text = Text3.Text
setting.Text3.Text = Text4.Text
SaveSetting "3", "id", "3", Text1.Text
SaveSetting "3", "pw", "3", Text2.Text
SaveSetting "3", "site", "3", Text3.Text
SaveSetting "3", "key", "3", Text4.Text
SaveSetting "ch", "3", "3", "che"
main.Check3.value = 1
main.Hide
MsgBox "Setting3 에관한 정보를 저장하였습니다", vbInformation, "알림"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End If

If Command14.Caption = "Setting4 저장" Then
main.Text7.Text = Text1.Text
main.Text8.Text = Text2.Text
main.Text30.Text = Text3.Text
setting.Text4.Text = Text4.Text
SaveSetting "4", "id", "3", Text1.Text
SaveSetting "4", "pw", "3", Text2.Text
SaveSetting "4", "site", "3", Text3.Text
SaveSetting "4", "key", "3", Text4.Text
SaveSetting "ch", "4", "3", "che4"
main.Check4.value = 1
main.Hide
MsgBox "Setting4 에관한 정보를 저장하였습니다", vbInformation, "알림"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End If

If Command14.Caption = "Setting5 저장" Then
main.Text17.Text = Text1.Text
main.Text16.Text = Text2.Text
main.Text34.Text = Text3.Text
setting.Text11.Text = Text4.Text
SaveSetting "5", "id", "3", Text1.Text
SaveSetting "5", "pw", "3", Text2.Text
SaveSetting "5", "site", "3", Text3.Text
SaveSetting "5", "key", "3", Text4.Text
SaveSetting "ch", "5", "3", "che5"
main.Check5.value = 1
main.Hide
MsgBox "Setting5 에관한 정보를 저장하였습니다", vbInformation, "알림"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End If

If Command14.Caption = "Setting6 저장" Then
main.Text15.Text = Text1.Text
main.Text14.Text = Text2.Text
main.Text35.Text = Text3.Text
setting.Text10.Text = Text4.Text
SaveSetting "6", "id", "3", Text1.Text
SaveSetting "6", "pw", "3", Text2.Text
SaveSetting "6", "site", "3", Text3.Text
SaveSetting "6", "key", "3", Text4.Text
SaveSetting "ch", "6", "3", "che6"
main.Check6.value = 1
main.Hide
MsgBox "Setting6 에관한 정보를 저장하였습니다", vbInformation, "알림"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End If

If Command14.Caption = "Setting7 저장" Then
main.Text13.Text = Text1.Text
main.Text12.Text = Text2.Text
main.Text36.Text = Text3.Text
setting.Text9.Text = Text4.Text
SaveSetting "7", "id", "3", Text1.Text
SaveSetting "7", "pw", "3", Text2.Text
SaveSetting "7", "site", "3", Text3.Text
SaveSetting "7", "key", "3", Text4.Text
SaveSetting "ch", "7", "3", "che7"
main.Check7.value = 1
main.Hide
MsgBox "Setting7 에관한 정보를 저장하였습니다", vbInformation, "알림"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End If

If Command14.Caption = "Setting8 저장" Then
main.Text11.Text = Text1.Text
main.Text10.Text = Text2.Text
main.Text37.Text = Text3.Text
setting.Text8.Text = Text4.Text
SaveSetting "8", "id", "3", Text1.Text
SaveSetting "8", "pw", "3", Text2.Text
SaveSetting "8", "site", "3", Text3.Text
SaveSetting "8", "key", "3", Text4.Text
SaveSetting "ch", "8", "3", "che8"
main.Check8.value = 1
main.Hide
MsgBox "Setting8 에관한 정보를 저장하였습니다", vbInformation, "알림"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End If

If Command14.Caption = "Setting9 저장" Then
main.Text18.Text = Text1.Text
main.Text19.Text = Text2.Text
main.Text26.Text = Text3.Text
setting.Text12.Text = Text4.Text
SaveSetting "9", "id", "3", Text1.Text
SaveSetting "9", "pw", "3", Text2.Text
SaveSetting "9", "site", "3", Text3.Text
SaveSetting "9", "key", "3", Text4.Text
SaveSetting "ch", "9", "3", "che9"
main.Check9.value = 1
main.Hide
MsgBox "Setting9 에관한 정보를 저장하였습니다", vbInformation, "알림"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End If

If Command14.Caption = "Setting10 저장" Then
main.Text20.Text = Text1.Text
main.Text21.Text = Text2.Text
main.Text27.Text = Text3.Text
setting.Text13.Text = Text4.Text
SaveSetting "10", "id", "3", Text1.Text
SaveSetting "10", "pw", "3", Text2.Text
SaveSetting "10", "site", "3", Text3.Text
SaveSetting "10", "key", "3", Text4.Text
SaveSetting "ch", "10", "3", "che10"
main.Check10.value = 1
main.Hide
MsgBox "Setting10 에관한 정보를 저장하였습니다", vbInformation, "알림"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End If

If Command14.Caption = "Setting11 저장" Then
main.Text22.Text = Text1.Text
main.Text23.Text = Text2.Text
main.Text28.Text = Text3.Text
setting.Text14.Text = Text4.Text
SaveSetting "11", "id", "3", Text1.Text
SaveSetting "11", "pw", "3", Text2.Text
SaveSetting "11", "site", "3", Text3.Text
SaveSetting "11", "key", "3", Text4.Text
SaveSetting "ch", "11", "3", "che11"
main.Check11.value = 1
main.Hide
MsgBox "Setting11 에관한 정보를 저장하였습니다", vbInformation, "알림"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End If

If Command14.Caption = "Setting12 저장" Then
main.Text24.Text = Text1.Text
main.Text25.Text = Text2.Text
main.Text29.Text = Text3.Text
setting.Text15.Text = Text4.Text
SaveSetting "12", "id", "3", Text1.Text
SaveSetting "12", "pw", "3", Text2.Text
SaveSetting "12", "site", "3", Text3.Text
SaveSetting "12", "key", "3", Text4.Text
SaveSetting "ch", "12", "3", "che12"
main.Check12.value = 1
main.Hide
MsgBox "Setting12 에관한 정보를 저장하였습니다", vbInformation, "알림"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End If
End Sub

Private Sub Command2_Click()
Command14.Caption = "Setting1 저장"
End Sub

Private Sub Command3_Click()
Command14.Caption = "Setting2 저장"
End Sub

Private Sub Command4_Click()
Command14.Caption = "Setting3 저장"
End Sub

Private Sub Command5_Click()
Command14.Caption = "Setting4 저장"
End Sub

Private Sub Command6_Click()
Command14.Caption = "Setting5 저장"
End Sub

Private Sub Command7_Click()
Command14.Caption = "Setting6 저장"
End Sub

Private Sub Command8_Click()
Command14.Caption = "Setting7 저장"
End Sub

Private Sub Command9_Click()
Command14.Caption = "Setting8 저장"
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.Show
SettingMagician.Hide
End Sub

Private Sub Text3_Click()
Text3.Text = ""
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
Text4 = KeyCode
End Sub
