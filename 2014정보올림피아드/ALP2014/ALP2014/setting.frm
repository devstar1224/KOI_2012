VERSION 5.00
Begin VB.Form setting 
   BorderStyle     =   0  '없음
   Caption         =   "setting"
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8550
   Icon            =   "setting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleMode       =   0  '사용자
   ScaleWidth      =   8535
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "자동입력 키설정"
      BeginProperty Font 
         Name            =   "a피오피네모OL"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   5280
      TabIndex        =   39
      Top             =   4800
      Width           =   4815
      Begin VB.CommandButton Command16 
         Caption         =   "기본값"
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
         Left            =   3120
         TabIndex        =   46
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton Command15 
         Caption         =   "저장"
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
         Left            =   3120
         TabIndex        =   45
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox Text15 
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Text14 
         BackColor       =   &H00FFC0C0&
         Height          =   270
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H00C0FFFF&
         Height          =   270
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H00C0FFC0&
         Height          =   270
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton Command13 
         Caption         =   "<"
         Height          =   375
         Left            =   3120
         TabIndex        =   40
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label32 
         BackStyle       =   0  '투명
         Caption         =   "해당하는 세팅을 누른후 설정하고싶은 키를 눌러주세요"
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   2040
         Width           =   4695
      End
      Begin VB.Label Label29 
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
         Left            =   1440
         TabIndex        =   66
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label28 
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
         Left            =   1440
         TabIndex        =   65
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label27 
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
         Left            =   1440
         TabIndex        =   64
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label26 
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
         Left            =   1440
         TabIndex        =   63
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label15 
         BackStyle       =   0  '투명
         Caption         =   "Setting 12:"
         BeginProperty Font 
            Name            =   "a피오피네모"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   50
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackStyle       =   0  '투명
         Caption         =   "Setting 11:"
         BeginProperty Font 
            Name            =   "a피오피네모"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   49
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackStyle       =   0  '투명
         Caption         =   "Setting 10:"
         BeginProperty Font 
            Name            =   "a피오피네모"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackStyle       =   0  '투명
         Caption         =   " Setting 9:"
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
         Left            =   120
         TabIndex        =   47
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "자동입력 키설정"
      BeginProperty Font 
         Name            =   "a피오피네모OL"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   25
      Top             =   4800
      Width           =   4815
      Begin VB.CommandButton Command12 
         Caption         =   "<"
         Height          =   375
         Left            =   3120
         TabIndex        =   38
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command11 
         Caption         =   ">"
         Height          =   375
         Left            =   3960
         TabIndex        =   37
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00C0FFC0&
         Height          =   270
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00C0FFFF&
         Height          =   270
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00FFC0C0&
         Height          =   270
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "기본값"
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
         Left            =   3120
         TabIndex        =   26
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton Command9 
         Caption         =   "저장"
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
         Left            =   3120
         TabIndex        =   27
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label31 
         BackStyle       =   0  '투명
         Caption         =   "해당하는 세팅을 누른후 설정하고싶은 키를 눌러주세요"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   2040
         Width           =   4695
      End
      Begin VB.Label Label25 
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
         Left            =   1440
         TabIndex        =   62
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label24 
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
         Left            =   1440
         TabIndex        =   61
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label23 
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
         Left            =   1440
         TabIndex        =   60
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label22 
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
         Left            =   1440
         TabIndex        =   59
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label11 
         BackStyle       =   0  '투명
         Caption         =   " Setting 5:"
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
         Left            =   120
         TabIndex        =   35
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackStyle       =   0  '투명
         Caption         =   "Setting 6:"
         BeginProperty Font 
            Name            =   "a피오피네모"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '투명
         Caption         =   "Setting 7:"
         BeginProperty Font 
            Name            =   "a피오피네모"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  '투명
         Caption         =   "Setting 8:"
         BeginProperty Font 
            Name            =   "a피오피네모"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   1560
         Width           =   1215
      End
   End
   Begin Project1.ctlSkin ctlSkin1 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   8281
      Caption         =   "//Setting// Automatic Login Program - ALP2014"
      Resize          =   0   'False
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "리스트 표시 시간설정"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   51
         Top             =   3840
         Width           =   4815
         Begin VB.Timer Timer1 
            Interval        =   1
            Left            =   960
            Top             =   240
         End
         Begin VB.TextBox Text16 
            Height          =   270
            Left            =   600
            TabIndex        =   52
            Text            =   "3000"
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label17 
            BackStyle       =   0  '투명
            Caption         =   "단축키:Shift + Tab"
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
            Left            =   2880
            TabIndex        =   54
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label16 
            BackStyle       =   0  '투명
            Caption         =   "1000 = 1초"
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
            Left            =   3240
            TabIndex        =   53
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "아이디 저장 레지스트리 초기화"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   5040
         TabIndex        =   23
         Top             =   3720
         Width           =   3375
         Begin VB.CommandButton Command6 
            Caption         =   "아이디,PW 초기화"
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
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   3135
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "프로그램 진입관련"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   5040
         TabIndex        =   13
         Top             =   480
         Width           =   3375
         Begin VB.CommandButton Command5 
            Caption         =   "초기화"
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
            Left            =   120
            TabIndex        =   22
            Top             =   2400
            Width           =   1335
         End
         Begin VB.CommandButton Command4 
            Caption         =   "저장"
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
            Left            =   1920
            TabIndex        =   21
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox Text7 
            Height          =   375
            Left            =   1320
            TabIndex        =   20
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox Text6 
            Height          =   375
            IMEMode         =   3  '사용 못함
            Left            =   840
            PasswordChar    =   "*"
            TabIndex        =   18
            Top             =   1320
            Width           =   2415
         End
         Begin VB.TextBox Text5 
            Height          =   375
            IMEMode         =   3  '사용 못함
            Left            =   840
            PasswordChar    =   "*"
            TabIndex        =   16
            Top             =   840
            Width           =   2415
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "프로그램 시작시 비밀번호 입력"
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
            Left            =   240
            TabIndex        =   14
            Top             =   480
            Width           =   3015
         End
         Begin VB.Label Label7 
            BackStyle       =   0  '투명
            Caption         =   "분실시 힌트:"
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
            Left            =   120
            TabIndex        =   19
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label Label6 
            BackStyle       =   0  '투명
            Caption         =   "재확인:"
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
            Left            =   120
            TabIndex        =   17
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label5 
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
            TabIndex        =   15
            Top             =   960
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "자동 시작 실행"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   11
         Top             =   3000
         Width           =   4815
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "프로그램 시작시 자동 시작"
            BeginProperty Font 
               Name            =   "a피오피네모"
               Size            =   12
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   12
            Top             =   360
            Width           =   4215
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "자동입력 키설정"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   4815
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   70
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton Command10 
            Caption         =   ">"
            Height          =   375
            Left            =   3960
            TabIndex        =   36
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton Command2 
            Caption         =   "기본값"
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
            Left            =   3120
            TabIndex        =   10
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "저장"
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
            Left            =   3120
            TabIndex        =   9
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H00E0E0E0&
            Height          =   270
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   1560
            Width           =   855
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H00FFC0C0&
            Height          =   270
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label30 
            BackStyle       =   0  '투명
            Caption         =   "해당하는 세팅을 누른후 설정하고싶은 키를 눌러주세요"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   2040
            Width           =   4695
         End
         Begin VB.Label Label21 
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
            Left            =   1320
            TabIndex        =   58
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label Label20 
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
            Left            =   1320
            TabIndex        =   57
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label19 
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
            Left            =   1320
            TabIndex        =   56
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label18 
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
            Left            =   1320
            TabIndex        =   55
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label4 
            BackStyle       =   0  '투명
            Caption         =   "Setting 4:"
            BeginProperty Font 
               Name            =   "a피오피네모"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '투명
            Caption         =   "Setting 3:"
            BeginProperty Font 
               Name            =   "a피오피네모"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '투명
            Caption         =   "Setting 2:"
            BeginProperty Font 
               Name            =   "a피오피네모"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   " Setting 1:"
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
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "setting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.value = 1 Then
SaveSetting "a", "u", "t", "az"
Else
DeleteSetting "a", "u", "t"
End If
End Sub

Private Sub Check2_Click()
If Check2.value = 1 Then
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Else
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
End If
End Sub

Private Sub Command1_Click()
Dim a
a = Date + Time
If MsgBox("정말로 저장을 하시겟습니까?", vbYesNo, "알림") = vbYes Then
SaveSetting "13", "key", "3", Text1.Text
SaveSetting "2", "key", "3", Text2.Text
SaveSetting "3", "key", "3", Text3.Text
SaveSetting "4", "key", "3", Text4.Text
SaveSetting "5", "key", "3", Text11.Text
SaveSetting "6", "key", "3", Text10.Text
SaveSetting "7", "key", "3", Text9.Text
SaveSetting "8", "key", "3", Text8.Text
SaveSetting "9", "key", "3", Text12.Text
SaveSetting "10", "key", "3", Text13.Text
SaveSetting "11", "key", "3", Text14.Text
SaveSetting "12", "key", "3", Text15.Text
SaveSetting "14", "key", "3", "q"
MsgBox "저장을 성공하였습니다.", vbInformation, "알림"
Else
MsgBox "사용자가 거절하여 저장을 실패하엿습니다", vbInformation, "알림"
End If

End Sub

Private Sub Command10_Click()
Frame5.Top = 480
Frame5.Left = 120
Frame1.Top = 4800
Frame1.Left = 5280
Frame6.Top = 4800
Frame6.Left = 5280
End Sub

Private Sub Command11_Click()
Frame6.Top = 480
Frame6.Left = 120
Frame1.Top = 4800
Frame1.Left = 5280
Frame5.Left = 5280
Frame5.Top = 4800
End Sub

Private Sub Command12_Click()
Frame6.Top = 4800
Frame6.Left = 5280
Frame5.Top = 4800
Frame5.Left = 5280
Frame1.Top = 480
Frame1.Left = 120
End Sub

Private Sub Command13_Click()
Frame6.Top = 4800
Frame6.Left = 5280
Frame5.Top = 480
Frame5.Left = 120
Frame1.Top = 4800
Frame1.Left = 5280
End Sub

Private Sub Command15_Click()
Dim a
a = Date + Time
If MsgBox("정말로 저장을 하시겟습니까?", vbYesNo, "알림") = vbYes Then

MsgBox "저장을 성공하였습니다.", vbInformation, "알림"
Else
MsgBox "사용자가 거절하여 저장을 실패하엿습니다", vbInformation, "알림"
End If
End Sub

Private Sub Command16_Click()
Dim a
a = Date + Time
If MsgBox("정말로 기본값으로 하시겟습니까?", vbYesNo, "알림") = vbYes Then
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text11.Text = ""
Text10.Text = ""
Text9.Text = ""
Text8.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
SaveSetting "13", "key", "3", Text1.Text
SaveSetting "2", "key", "3", Text2.Text
SaveSetting "3", "key", "3", Text3.Text
SaveSetting "4", "key", "3", Text4.Text
SaveSetting "5", "key", "3", Text11.Text
SaveSetting "6", "key", "3", Text10.Text
SaveSetting "7", "key", "3", Text9.Text
SaveSetting "8", "key", "3", Text8.Text
SaveSetting "9", "key", "3", Text12.Text
SaveSetting "10", "key", "3", Text13.Text
SaveSetting "11", "key", "3", Text14.Text
SaveSetting "12", "key", "3", Text15.Text
SaveSetting "14", "key", "3", "q"
MsgBox "성공하였습니다.", vbInformation, "알림"
Else
MsgBox "사용자가 거절하여 기본값으로 설정을 실패하엿습니다", vbInformation, "알림"
End If
End Sub



Private Sub Command2_Click()
Dim a
a = Date + Time
If MsgBox("정말로 기본값으로 하시겟습니까?", vbYesNo, "알림") = vbYes Then
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text11.Text = ""
Text10.Text = ""
Text9.Text = ""
Text8.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
SaveSetting "13", "key", "3", Text1.Text
SaveSetting "2", "key", "3", Text2.Text
SaveSetting "3", "key", "3", Text3.Text
SaveSetting "4", "key", "3", Text4.Text
SaveSetting "5", "key", "3", Text11.Text
SaveSetting "6", "key", "3", Text10.Text
SaveSetting "7", "key", "3", Text9.Text
SaveSetting "8", "key", "3", Text8.Text
SaveSetting "9", "key", "3", Text12.Text
SaveSetting "10", "key", "3", Text13.Text
SaveSetting "11", "key", "3", Text14.Text
SaveSetting "12", "key", "3", Text15.Text
SaveSetting "14", "key", "3", "q"
MsgBox "성공하였습니다.", vbInformation, "알림"
Else
MsgBox "사용자가 거절하여 기본값으로 설정을 실패하엿습니다", vbInformation, "알림"
End If

End Sub


Private Sub Command4_Click()
If Text5.Text = Text6.Text Then
GoTo g
Else
MsgBox "비밀번호가 일치하지 않습니다", vbInformation, "안내"
GoTo q
End If
g:
If MsgBox("저장을 하시겟습니까?", vbYesNo, "알림") = vbYes Then
SaveSetting "14", "key", "3", "qq"
SaveSetting "1", "3", "pw", Text5.Text
SaveSetting "1", "2", "hint", Text7.Text
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text7.Text = "이미 설정됨"
Command4.Enabled = False
Check2.Enabled = False
MsgBox "저장되었습니다", vbInformation, "안내"
Else
MsgBox "취소 되었습니다", vbInformation, "알림"
End If
q:
End Sub

Private Sub Command5_Click()
de.Show
Unload Me
End Sub

Private Sub Command6_Click()
Dim a
a = Date + Time
If MsgBox("정말로 초기화를 하시겟습니까?", vbYesNo, "알림") = vbYes Then
DeleteSetting "1", "id", "3"
DeleteSetting "1", "pw", "3"
DeleteSetting "2", "id", "3"
DeleteSetting "2", "pw", "3"
DeleteSetting "3", "id", "3"
DeleteSetting "3", "pw", "3"
DeleteSetting "4", "id", "3"
DeleteSetting "4", "pw", "3"
DeleteSetting "5", "id", "3"
DeleteSetting "5", "pw", "3"
DeleteSetting "6", "id", "3"
DeleteSetting "6", "pw", "3"
DeleteSetting "7", "id", "3"
DeleteSetting "7", "pw", "3"
DeleteSetting "8", "id", "3"
DeleteSetting "8", "pw", "3"
DeleteSetting "9", "id", "3"
DeleteSetting "9", "pw", "3"
DeleteSetting "10", "id", "3"
DeleteSetting "10", "pw", "3"
DeleteSetting "11", "id", "3"
DeleteSetting "11", "pw", "3"
DeleteSetting "12", "id", "3"
DeleteSetting "12", "pw", "3"
DeleteSetting "1", "site", "3"
DeleteSetting "2", "site", "3"
DeleteSetting "3", "site", "3"
DeleteSetting "4", "site", "3"
DeleteSetting "5", "site", "3"
DeleteSetting "6", "site", "3"
DeleteSetting "7", "site", "3"
DeleteSetting "8", "site", "3"
DeleteSetting "9", "site", "3"
DeleteSetting "10", "site", "3"
DeleteSetting "11", "site", "3"
DeleteSetting "12", "site", "3"

MsgBox "초기화를 성공하였습니다.", vbInformation, "알림"
Else

MsgBox "사용자가 거절하여 초기화를 실패하엿습니다", vbInformation, "알림"
End If
End Sub


Private Sub Command8_Click()
Dim a
a = Date + Time
If MsgBox("정말로 기본값으로 하시겟습니까?", vbYesNo, "알림") = vbYes Then
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text11.Text = ""
Text10.Text = ""
Text9.Text = ""
Text8.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
SaveSetting "1", "key", "3", Text1.Text
SaveSetting "2", "key", "3", Text2.Text
SaveSetting "3", "key", "3", Text3.Text
SaveSetting "4", "key", "3", Text4.Text
SaveSetting "5", "key", "3", Text11.Text
SaveSetting "6", "key", "3", Text10.Text
SaveSetting "7", "key", "3", Text9.Text
SaveSetting "8", "key", "3", Text8.Text
SaveSetting "9", "key", "3", Text12.Text
SaveSetting "10", "key", "3", Text13.Text
SaveSetting "11", "key", "3", Text14.Text
SaveSetting "12", "key", "3", Text15.Text
SaveSetting "14", "key", "3", "q"
MsgBox "성공하였습니다.", vbInformation, "알림"
Else
MsgBox "사용자가 거절하여 기본값으로 설정을 실패하엿습니다", vbInformation, "알림"
End If
End Sub

Private Sub Command9_Click()
Dim a
a = Date + Time
If MsgBox("정말로 저장을 하시겟습니까?", vbYesNo, "알림") = vbYes Then
SaveSetting "13", "key", "3", Text1.Text
SaveSetting "2", "key", "3", Text2.Text
SaveSetting "3", "key", "3", Text3.Text
SaveSetting "4", "key", "3", Text4.Text
SaveSetting "5", "key", "3", Text11.Text
SaveSetting "6", "key", "3", Text10.Text
SaveSetting "7", "key", "3", Text9.Text
SaveSetting "8", "key", "3", Text8.Text
SaveSetting "9", "key", "3", Text12.Text
SaveSetting "10", "key", "3", Text13.Text
SaveSetting "11", "key", "3", Text14.Text
SaveSetting "12", "key", "3", Text15.Text
SaveSetting "14", "key", "3", "q"
MsgBox "저장을 성공하였습니다.", vbInformation, "알림"
Else
MsgBox "사용자가 거절하여 저장을 실패하엿습니다", vbInformation, "알림"
End If
End Sub

Private Sub Form_Load()
Text1.Text = GetSetting("13", "key", "3")
Text2.Text = GetSetting("2", "key", "3")
Text3.Text = GetSetting("3", "key", "3")
Text4.Text = GetSetting("4", "key", "3")
Text11.Text = GetSetting("5", "key", "3")
Text10.Text = GetSetting("6", "key", "3")
Text9.Text = GetSetting("7", "key", "3")
Text8.Text = GetSetting("8", "key", "3")
Text12.Text = GetSetting("9", "key", "3")
Text13.Text = GetSetting("10", "key", "3")
Text14.Text = GetSetting("11", "key", "3")
Text15.Text = GetSetting("12", "key", "3")
 
If GetSetting("14", "key", "3") = "qq" Then
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text7.Text = "이미 설정됨"
Command4.Enabled = False
Check2.Enabled = False
Else
Check2.value = 0
End If
If Text16.Text = 3000 Then
Else
Text16.Text = GetSetting("list", "time", "3")
End If
Command4.Enabled = False
If Text7.Text = "이미 설정됨" Then
Command5.Enabled = True
Else
Command5.Enabled = False
End If
If GetSetting("a", "u", "t") = "az" Then
Check1.value = 1
Else
Check1.value = 0
End If

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Text1 = KeyCode
End Sub
Private Sub Text10_KeyDown(KeyCode As Integer, Shift As Integer)
Text10 = KeyCode
End Sub
Private Sub Text11_KeyDown(KeyCode As Integer, Shift As Integer)
Text11 = KeyCode
End Sub
Private Sub Text12_KeyDown(KeyCode As Integer, Shift As Integer)
Text12 = KeyCode
End Sub
Private Sub Text13_KeyDown(KeyCode As Integer, Shift As Integer)
Text13 = KeyCode
End Sub
Private Sub Text14_KeyDown(KeyCode As Integer, Shift As Integer)
Text14 = KeyCode
End Sub
Private Sub Text15_KeyDown(KeyCode As Integer, Shift As Integer)
Text15 = KeyCode
End Sub
Private Sub Text16_Change()
If IsNumeric(Text16.Text) = True Then
Else
Text16.Text = "3000"
MsgBox "숫자만 입력해주세요", vbInformation, "알림"
End If
End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
Text2 = KeyCode
End Sub
Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
Text3 = KeyCode
End Sub
Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
Text4 = KeyCode
End Sub
Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
Text8 = KeyCode
End Sub
Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
Text9 = KeyCode
End Sub
Private Sub Timer1_Timer()
SaveSetting "list", "time", "3", Text16.Text
End Sub
