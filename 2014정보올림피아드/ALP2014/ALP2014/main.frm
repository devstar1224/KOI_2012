VERSION 5.00
Begin VB.Form main 
   BorderStyle     =   0  '없음
   Caption         =   "main"
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8220
   BeginProperty Font 
      Name            =   "a피오피네모OL"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   8220
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ID Setting"
      BeginProperty Font 
         Name            =   "a피오피네모OL"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   1200
      TabIndex        =   66
      Top             =   6480
      Width           =   5775
      Begin VB.TextBox Text29 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   116
         Top             =   3000
         Width           =   1695
      End
      Begin VB.TextBox Text28 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   114
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox Text27 
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   112
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox Text26 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   110
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Timer Timer24 
         Enabled         =   0   'False
         Interval        =   90
         Left            =   3480
         Top             =   2520
      End
      Begin VB.Timer Timer23 
         Enabled         =   0   'False
         Interval        =   90
         Left            =   3960
         Top             =   2520
      End
      Begin VB.Timer Timer22 
         Enabled         =   0   'False
         Interval        =   90
         Left            =   600
         Top             =   2520
      End
      Begin VB.Timer Timer21 
         Enabled         =   0   'False
         Interval        =   90
         Left            =   1080
         Top             =   2520
      End
      Begin VB.Timer Timer20 
         Enabled         =   0   'False
         Interval        =   90
         Left            =   3120
         Top             =   1080
      End
      Begin VB.Timer Timer19 
         Enabled         =   0   'False
         Interval        =   90
         Left            =   3600
         Top             =   1080
      End
      Begin VB.Timer Timer18 
         Enabled         =   0   'False
         Interval        =   90
         Left            =   1440
         Top             =   1080
      End
      Begin VB.Timer Timer17 
         Enabled         =   0   'False
         Interval        =   90
         Left            =   1920
         Top             =   1080
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "사용여부 체크"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   77
         Top             =   3360
         Width           =   5535
         Begin VB.CheckBox Check12 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Setting 12"
            BeginProperty Font 
               Name            =   "a피오피네모"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   97
            Top             =   360
            Width           =   1695
         End
         Begin VB.CheckBox Check11 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Setting 11"
            BeginProperty Font 
               Name            =   "a피오피네모"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2880
            TabIndex        =   96
            Top             =   360
            Width           =   1695
         End
         Begin VB.CheckBox Check10 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Setting 10"
            BeginProperty Font 
               Name            =   "a피오피네모"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1560
            TabIndex        =   95
            Top             =   360
            Width           =   1695
         End
         Begin VB.CheckBox Check9 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Setting 9"
            BeginProperty Font 
               Name            =   "a피오피네모"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   94
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "a피오피네모"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   76
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text25 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  '사용 못함
         Left            =   3000
         PasswordChar    =   "*"
         TabIndex        =   75
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox Text24 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3000
         TabIndex        =   74
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox Text23 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  '사용 못함
         Left            =   600
         PasswordChar    =   "*"
         TabIndex        =   73
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox Text22 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         TabIndex        =   72
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox Text21 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  '사용 못함
         Left            =   3000
         PasswordChar    =   "*"
         TabIndex        =   71
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox Text20 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3000
         TabIndex        =   70
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox Text19 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  '사용 못함
         Left            =   600
         PasswordChar    =   "*"
         TabIndex        =   69
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox Text18 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   600
         TabIndex        =   68
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton Command12 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   67
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label51 
         BackStyle       =   0  '투명
         Caption         =   "site:"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   115
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label50 
         BackStyle       =   0  '투명
         Caption         =   "site:"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   113
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label49 
         BackStyle       =   0  '투명
         Caption         =   "site:"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   111
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label48 
         BackStyle       =   0  '투명
         Caption         =   "site:"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   109
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label39 
         BackStyle       =   0  '투명
         Caption         =   "PW:"
         Height          =   255
         Left            =   2640
         TabIndex        =   89
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label38 
         BackStyle       =   0  '투명
         Caption         =   "ID:"
         Height          =   375
         Left            =   2640
         TabIndex        =   88
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label37 
         BackStyle       =   0  '투명
         Caption         =   "ID Setting 12"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   87
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label36 
         BackStyle       =   0  '투명
         Caption         =   "PW:"
         Height          =   375
         Left            =   240
         TabIndex        =   86
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label35 
         BackStyle       =   0  '투명
         Caption         =   "ID:"
         Height          =   375
         Left            =   240
         TabIndex        =   85
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label34 
         BackStyle       =   0  '투명
         Caption         =   "ID Setting 11"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   84
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label33 
         BackStyle       =   0  '투명
         Caption         =   "PW:"
         Height          =   375
         Left            =   2640
         TabIndex        =   83
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label32 
         BackStyle       =   0  '투명
         Caption         =   "ID:"
         Height          =   255
         Left            =   2640
         TabIndex        =   82
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label31 
         BackStyle       =   0  '투명
         Caption         =   "ID Setting 10"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   81
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label30 
         BackStyle       =   0  '투명
         Caption         =   "PW:"
         Height          =   375
         Left            =   240
         TabIndex        =   80
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label29 
         BackStyle       =   0  '투명
         Caption         =   "ID:"
         Height          =   375
         Left            =   240
         TabIndex        =   79
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label28 
         BackStyle       =   0  '투명
         Caption         =   "ID setting 9"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   78
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ID Setting"
      BeginProperty Font 
         Name            =   "a피오피네모OL"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   8520
      TabIndex        =   41
      Top             =   5040
      Width           =   5775
      Begin VB.TextBox Text37 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  '사용 못함
         Left            =   3240
         TabIndex        =   132
         Top             =   3000
         Width           =   1695
      End
      Begin VB.TextBox Text36 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  '사용 못함
         Left            =   720
         TabIndex        =   129
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox Text35 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  '사용 못함
         Left            =   3240
         TabIndex        =   127
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox Text34 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  '사용 못함
         Left            =   720
         TabIndex        =   126
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Timer Timer16 
         Enabled         =   0   'False
         Interval        =   90
         Left            =   3480
         Top             =   2520
      End
      Begin VB.Timer Timer15 
         Enabled         =   0   'False
         Interval        =   90
         Left            =   3960
         Top             =   2520
      End
      Begin VB.Timer Timer14 
         Enabled         =   0   'False
         Interval        =   90
         Left            =   600
         Top             =   2400
      End
      Begin VB.Timer Timer13 
         Enabled         =   0   'False
         Interval        =   90
         Left            =   1080
         Top             =   2400
      End
      Begin VB.Timer Timer12 
         Enabled         =   0   'False
         Interval        =   90
         Left            =   3360
         Top             =   1200
      End
      Begin VB.Timer Timer11 
         Enabled         =   0   'False
         Interval        =   90
         Left            =   3720
         Top             =   1200
      End
      Begin VB.Timer Timer10 
         Enabled         =   0   'False
         Interval        =   90
         Left            =   1200
         Top             =   960
      End
      Begin VB.Timer Timer9 
         Enabled         =   0   'False
         Interval        =   90
         Left            =   1680
         Top             =   960
      End
      Begin VB.CommandButton Command11 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   65
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Text17 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   600
         TabIndex        =   52
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox Text16 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  '사용 못함
         Left            =   600
         PasswordChar    =   "*"
         TabIndex        =   51
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox Text15 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3000
         TabIndex        =   50
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox Text14 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  '사용 못함
         Left            =   3000
         PasswordChar    =   "*"
         TabIndex        =   49
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         TabIndex        =   48
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  '사용 못함
         Left            =   600
         PasswordChar    =   "*"
         TabIndex        =   47
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3000
         TabIndex        =   46
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  '사용 못함
         Left            =   3000
         PasswordChar    =   "*"
         TabIndex        =   45
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "a피오피네모"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   44
         Top             =   360
         Width           =   735
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "사용여부 체크"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   43
         Top             =   3360
         Width           =   5535
         Begin VB.CheckBox Check8 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Setting 8"
            BeginProperty Font 
               Name            =   "a피오피네모"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4080
            TabIndex        =   93
            Top             =   360
            Width           =   1335
         End
         Begin VB.CheckBox Check7 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Setting 7"
            BeginProperty Font 
               Name            =   "a피오피네모"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   92
            Top             =   360
            Width           =   1695
         End
         Begin VB.CheckBox Check6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Setting 6"
            BeginProperty Font 
               Name            =   "a피오피네모"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   91
            Top             =   360
            Width           =   1695
         End
         Begin VB.CheckBox Check5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Setting 5"
            BeginProperty Font 
               Name            =   "a피오피네모"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.CommandButton Command7 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   42
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label59 
         BackStyle       =   0  '투명
         Caption         =   "site:"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   131
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label58 
         BackStyle       =   0  '투명
         Caption         =   "site:"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   130
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label57 
         BackStyle       =   0  '투명
         Caption         =   "site:"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   128
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label56 
         BackStyle       =   0  '투명
         Caption         =   "site:"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   125
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label27 
         BackStyle       =   0  '투명
         Caption         =   "ID setting 5"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   64
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label26 
         BackStyle       =   0  '투명
         Caption         =   "ID:"
         Height          =   375
         Left            =   240
         TabIndex        =   63
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label25 
         BackStyle       =   0  '투명
         Caption         =   "PW:"
         Height          =   375
         Left            =   240
         TabIndex        =   62
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label24 
         BackStyle       =   0  '투명
         Caption         =   "ID Setting 6"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   61
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label23 
         BackStyle       =   0  '투명
         Caption         =   "ID:"
         Height          =   255
         Left            =   2640
         TabIndex        =   60
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label22 
         BackStyle       =   0  '투명
         Caption         =   "PW:"
         Height          =   375
         Left            =   2640
         TabIndex        =   59
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label21 
         BackStyle       =   0  '투명
         Caption         =   "ID Setting 7"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   58
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label20 
         BackStyle       =   0  '투명
         Caption         =   "ID:"
         Height          =   375
         Left            =   240
         TabIndex        =   57
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label19 
         BackStyle       =   0  '투명
         Caption         =   "PW:"
         Height          =   375
         Left            =   240
         TabIndex        =   56
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label18 
         BackStyle       =   0  '투명
         Caption         =   "ID Setting 8"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   55
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label17 
         BackStyle       =   0  '투명
         Caption         =   "ID:"
         Height          =   375
         Left            =   2640
         TabIndex        =   54
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label16 
         BackStyle       =   0  '투명
         Caption         =   "PW:"
         Height          =   255
         Left            =   2640
         TabIndex        =   53
         Top             =   2760
         Width           =   375
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ID Setting"
      BeginProperty Font 
         Name            =   "a피오피네모OL"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   8520
      TabIndex        =   8
      Top             =   480
      Width           =   5775
      Begin VB.TextBox Text33 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  '사용 못함
         Left            =   3240
         TabIndex        =   120
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox Text32 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  '사용 못함
         Left            =   720
         TabIndex        =   119
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox Text31 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  '사용 못함
         Left            =   720
         TabIndex        =   118
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox Text30 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  '사용 못함
         Left            =   3240
         TabIndex        =   117
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Timer Timer8 
         Enabled         =   0   'False
         Interval        =   90
         Left            =   3120
         Top             =   2400
      End
      Begin VB.Timer Timer7 
         Enabled         =   0   'False
         Interval        =   90
         Left            =   3600
         Top             =   2400
      End
      Begin VB.Timer Timer6 
         Enabled         =   0   'False
         Interval        =   90
         Left            =   480
         Top             =   2280
      End
      Begin VB.Timer Timer5 
         Enabled         =   0   'False
         Interval        =   90
         Left            =   1080
         Top             =   2280
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   90
         Left            =   3240
         Top             =   960
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   90
         Left            =   3720
         Top             =   960
      End
      Begin VB.CommandButton Command5 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   40
         Top             =   1920
         Width           =   615
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "사용여부 체크"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   31
         Top             =   3360
         Width           =   5535
         Begin VB.CheckBox Check3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Setting 3"
            BeginProperty Font 
               Name            =   "a피오피네모"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   34
            Top             =   360
            Width           =   1335
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Setting 2"
            BeginProperty Font 
               Name            =   "a피오피네모"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   33
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Setting 1"
            BeginProperty Font 
               Name            =   "a피오피네모"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   360
            Width           =   1695
         End
         Begin VB.CheckBox Check4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Setting4"
            BeginProperty Font 
               Name            =   "a피오피네모"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4080
            TabIndex        =   35
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "a피오피네모"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   36
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  '사용 못함
         Left            =   3120
         PasswordChar    =   "*"
         TabIndex        =   30
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   28
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  '사용 못함
         Left            =   600
         PasswordChar    =   "*"
         TabIndex        =   25
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         TabIndex        =   23
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  '사용 못함
         Left            =   3000
         PasswordChar    =   "*"
         TabIndex        =   20
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3000
         TabIndex        =   19
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  '사용 못함
         Left            =   600
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   600
         TabIndex        =   11
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label55 
         BackStyle       =   0  '투명
         Caption         =   "site:"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   124
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label54 
         BackStyle       =   0  '투명
         Caption         =   "site:"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   123
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label53 
         BackStyle       =   0  '투명
         Caption         =   "site:"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   122
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label52 
         BackStyle       =   0  '투명
         Caption         =   "site:"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   121
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label15 
         BackStyle       =   0  '투명
         Caption         =   "PW:"
         Height          =   255
         Left            =   2760
         TabIndex        =   29
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label14 
         BackStyle       =   0  '투명
         Caption         =   "ID:"
         Height          =   375
         Left            =   2760
         TabIndex        =   27
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label13 
         BackStyle       =   0  '투명
         Caption         =   "ID Setting 4"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   26
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label12 
         BackStyle       =   0  '투명
         Caption         =   "PW:"
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label11 
         BackStyle       =   0  '투명
         Caption         =   "ID:"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label10 
         BackStyle       =   0  '투명
         Caption         =   "ID Setting 3"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   21
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '투명
         Caption         =   "PW:"
         Height          =   375
         Left            =   2640
         TabIndex        =   18
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label8 
         BackStyle       =   0  '투명
         Caption         =   "ID:"
         Height          =   255
         Left            =   2640
         TabIndex        =   17
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label7 
         BackStyle       =   0  '투명
         Caption         =   "ID Setting 2"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   16
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '투명
         Caption         =   "PW:"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '투명
         Caption         =   "ID:"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "ID setting 1"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   13
         Top             =   480
         Width           =   1695
      End
   End
   Begin Project1.ctlSkin ctlSkin1 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   11245
      Caption         =   "//Main// Automatic Login Program - ALP2014"
      Resize          =   0   'False
      Begin VB.Frame Frame3 
         BackColor       =   &H80000005&
         Caption         =   "Log"
         BeginProperty Font 
            Name            =   "@a피오피네모OL"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         TabIndex        =   7
         Top             =   4680
         Width           =   7935
         Begin VB.TextBox Text9 
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            OLEDragMode     =   1  '자동
            ScrollBars      =   2  '수직
            TabIndex        =   37
            Top             =   360
            Width           =   7695
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000005&
         Caption         =   "Menu"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4215
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   2055
         Begin VB.CommandButton Command14 
            Caption         =   "아이디 리스트"
            BeginProperty Font 
               Name            =   "a피오피네모"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   106
            Top             =   3360
            Width           =   1815
         End
         Begin VB.CommandButton Command1 
            Caption         =   "프로그램 정보"
            BeginProperty Font 
               Name            =   "a피오피네모"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   2
            Top             =   2760
            Width           =   1815
         End
         Begin VB.CommandButton Command15 
            Caption         =   "설정마법사"
            BeginProperty Font 
               Name            =   "a피오피네모"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   840
            TabIndex        =   108
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Timer start 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   1080
            Top             =   1440
         End
         Begin VB.Timer stop1 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   600
            Top             =   1440
         End
         Begin VB.Timer Timer2 
            Enabled         =   0   'False
            Interval        =   180
            Left            =   240
            Top             =   240
         End
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   90
            Left            =   720
            Top             =   240
         End
         Begin VB.CommandButton Command2 
            Caption         =   "설정"
            BeginProperty Font 
               Name            =   "a피오피네모"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   3
            Top             =   2160
            Width           =   735
         End
         Begin VB.CommandButton Command3 
            Caption         =   "업데이트 확인"
            BeginProperty Font 
               Name            =   "a피오피네모"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   4
            Top             =   1560
            Width           =   1815
         End
         Begin VB.CommandButton Command4 
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
            Height          =   615
            Left            =   120
            TabIndex        =   5
            Top             =   960
            Width           =   1815
         End
         Begin VB.CommandButton Command10 
            Caption         =   "시작"
            BeginProperty Font 
               Name            =   "a피오피네모"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1080
            TabIndex        =   39
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton Command9 
            Caption         =   "중지"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "a피오피네모"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000005&
         Caption         =   "Main"
         BeginProperty Font 
            Name            =   "a피오피네모OL"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4215
         Left            =   2280
         TabIndex        =   6
         Top             =   480
         Width           =   5775
         Begin VB.Timer Timer25 
            Interval        =   1
            Left            =   480
            Top             =   1440
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '투명
            Caption         =   "Ver:1.0.0"
            BeginProperty Font 
               Name            =   "a피오피네모OL"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2280
            TabIndex        =   107
            Top             =   3720
            Width           =   2415
         End
         Begin VB.Label Label47 
            BackStyle       =   0  '투명
            Caption         =   "자동 로그인 프로그램(핫키형식)"
            BeginProperty Font 
               Name            =   "a피오피네모OL"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1320
            TabIndex        =   105
            Top             =   960
            Width           =   4815
         End
         Begin VB.Label Label46 
            BackStyle       =   0  '투명
            Caption         =   "(ALP2014)"
            BeginProperty Font 
               Name            =   "a피오피네모OL"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2040
            TabIndex        =   104
            Top             =   1320
            Width           =   1935
         End
         Begin VB.Label Label45 
            BackStyle       =   0  '투명
            Caption         =   "rogram"
            BeginProperty Font 
               Name            =   "a피오피네모OL"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4080
            TabIndex        =   103
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label44 
            BackStyle       =   0  '투명
            Caption         =   "P"
            BeginProperty Font 
               Name            =   "a피오피네모OL"
               Size            =   21.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3720
            TabIndex        =   102
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label42 
            BackStyle       =   0  '투명
            Caption         =   "ogin"
            BeginProperty Font 
               Name            =   "a피오피네모OL"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   101
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label43 
            BackStyle       =   0  '투명
            Caption         =   "L"
            BeginProperty Font 
               Name            =   "a피오피네모OL"
               Size            =   21.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2520
            TabIndex        =   100
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label41 
            BackStyle       =   0  '투명
            Caption         =   "utomatic"
            BeginProperty Font 
               Name            =   "a피오피네모OL"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   960
            TabIndex        =   99
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label40 
            BackStyle       =   0  '투명
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "a피오피네모OL"
               Size            =   21.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   600
            TabIndex        =   98
            Top             =   480
            Width           =   1935
         End
         Begin VB.Line Line4 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            X1              =   2400
            X2              =   3240
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Line Line3 
            BorderWidth     =   3
            X1              =   4320
            X2              =   4560
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Line Line2 
            BorderWidth     =   3
            X1              =   4320
            X2              =   4560
            Y1              =   2280
            Y2              =   2160
         End
         Begin VB.Line Line1 
            BorderWidth     =   3
            Index           =   0
            X1              =   4200
            X2              =   4320
            Y1              =   2160
            Y2              =   1920
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '투명
            Caption         =   "Menu를 클릭해주세요."
            BeginProperty Font 
               Name            =   "a피오피네모OL"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   720
            TabIndex        =   10
            Top             =   2880
            Width           =   4335
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "왼쪽에 해당하는"
            BeginProperty Font 
               Name            =   "a피오피네모OL"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   1320
            TabIndex        =   9
            Top             =   2160
            Width           =   3255
         End
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sw As Integer
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer


Private Sub Check10_Click()
Dim a
a = Date + Time
If Check10.value = 1 Then
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting - Setting10 Check Enable" & vbCrLf
SaveSetting "ch", "10", "3", "che10"
Text20.Enabled = True
Text21.Enabled = True
Text20.Text = ""
Text21.Text = ""
Text27.Text = ""
setting.Text10.Text = ""
Else
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting - Setting10 Check Disable" & vbCrLf
DeleteSetting "ch", "10", "3"
Text20.Text = ""
Text21.Text = ""
Text27.Text = ""
Text20.Enabled = False
Text21.Enabled = False
setting.Text10.Text = ""
End If
End Sub
Private Sub Check11_Click()
Dim a
a = Date + Time
If Check11.value = 1 Then
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting - Setting11 Check Enable" & vbCrLf
SaveSetting "ch", "11", "3", "che11"
Text22.Enabled = True
Text23.Enabled = True
Text22.Text = ""
Text23.Text = ""
Text28.Text = ""
setting.Text11.Text = ""
Else
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting - Setting11 Check Disable" & vbCrLf
DeleteSetting "ch", "11", "3"
Text22.Text = ""
Text23.Text = ""
Text28.Text = ""
Text22.Enabled = False
Text23.Enabled = False
setting.Text11.Text = ""
End If
End Sub

Private Sub Check12_Click()
Dim a
a = Date + Time
If Check12.value = 1 Then
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting - Setting12 Check Enable" & vbCrLf
SaveSetting "ch", "12", "3", "che12"
Text24.Enabled = True
Text25.Enabled = True
Text24.Text = ""
Text25.Text = ""
Text29.Text = ""
setting.Text12.Text = ""
Else
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting - Setting12 Check Disable" & vbCrLf
DeleteSetting "ch", "12", "3"
Text24.Text = ""
Text25.Text = ""
Text29.Text = ""
Text24.Enabled = False
Text25.Enabled = False
setting.Text12.Text = ""
End If
End Sub

Private Sub Check5_Click()
Dim a
a = Date + Time
If Check5.value = 1 Then
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting - Setting5 Check Enable" & vbCrLf
SaveSetting "ch", "5", "3", "che5"
Text17.Enabled = True
Text16.Enabled = True
Text17.Text = ""
Text16.Text = ""
Text34.Text = ""
setting.Text5.Text = ""
Else
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting - Setting5 Check Disable" & vbCrLf
DeleteSetting "ch", "5", "3"
Text17.Text = ""
Text16.Text = ""
Text34.Text = ""
Text17.Enabled = False
Text16.Enabled = False
setting.Text5.Text = ""
End If
End Sub

Private Sub Check6_Click()
Dim a
a = Date + Time
If Check6.value = 1 Then
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting - Setting6 Check Enable" & vbCrLf
SaveSetting "ch", "6", "3", "che6"
Text15.Enabled = True
Text14.Enabled = True
Text15.Text = ""
Text14.Text = ""
Text35.Text = ""
setting.Text6.Text = ""
Else
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting - Setting6 Check Disable" & vbCrLf
DeleteSetting "ch", "6", "3"
Text15.Text = ""
Text14.Text = ""
Text35.Text = ""
Text15.Enabled = False
Text14.Enabled = False
setting.Text6.Text = ""
End If
End Sub

Private Sub Check7_Click()
Dim a
a = Date + Time
If Check7.value = 1 Then
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting - Setting7 Check Enable" & vbCrLf
SaveSetting "ch", "7", "3", "che7"
Text13.Enabled = True
Text12.Enabled = True
Text13.Text = ""
Text12.Text = ""
Text36.Text = ""
setting.Text7.Text = ""
Else
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting - Setting7 Check Disable" & vbCrLf
DeleteSetting "ch", "7", "3"
Text13.Text = ""
Text12.Text = ""
Text36.Text = ""
Text13.Enabled = False
Text12.Enabled = False
setting.Text7.Text = ""
End If
End Sub

Private Sub Check8_Click()
Dim a
a = Date + Time
If Check8.value = 1 Then
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting - Setting8 Check Enable" & vbCrLf
SaveSetting "ch", "8", "3", "che8"
Text11.Enabled = True
Text10.Enabled = True
Text11.Text = ""
Text10.Text = ""
Text37.Text = ""
setting.Text8.Text = ""
Else
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting - Setting8 Check Disable" & vbCrLf
DeleteSetting "ch", "8", "3"
Text11.Text = ""
Text10.Text = ""
Text37.Text = ""
Text11.Enabled = False
Text10.Enabled = False
setting.Text8.Text = ""
End If
End Sub

Private Sub Check9_Click()
Dim a
a = Date + Time
If Check9.value = 1 Then
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting - Setting9 Check Enable" & vbCrLf
SaveSetting "ch", "9", "3", "che9"
Text18.Enabled = True
Text19.Enabled = True
Text18.Text = ""
Text19.Text = ""
Text26.Text = ""
setting.Text9.Text = ""
Else
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting - Setting9 Check Disable" & vbCrLf
DeleteSetting "ch", "9", "3"
Text18.Text = ""
Text19.Text = ""
Text26.Text = ""
Text18.Enabled = False
Text19.Enabled = False
setting.Text9.Text = ""
End If
End Sub
Private Sub Command11_Click()
Frame6.Left = 8520
Frame6.Top = 480
Frame4.Left = 2280
Frame4.Top = 480
Frame8.Left = 8520
Frame8.Top = 480
End Sub
Private Sub Command12_Click()
Frame6.Left = 2280
Frame6.Top = 480
Frame4.Left = 8520
Frame4.Top = 480
Frame8.Left = 8520
Frame8.Top = 480
End Sub
Private Sub Command13_Click()
Dim a
a = Date + Time
If MsgBox("정말로 저장을 하시겟습니까?", vbYesNo, "알림") = vbYes Then
SaveSetting "1", "id", "3", Text1.Text
SaveSetting "1", "pw", "3", Text2.Text
SaveSetting "2", "id", "3", Text3.Text
SaveSetting "2", "pw", "3", Text4.Text
SaveSetting "3", "id", "3", Text5.Text
SaveSetting "3", "pw", "3", Text6.Text
SaveSetting "4", "id", "3", Text7.Text
SaveSetting "4", "pw", "3", Text8.Text
SaveSetting "5", "id", "3", Text17.Text
SaveSetting "5", "pw", "3", Text16.Text
SaveSetting "6", "id", "3", Text15.Text
SaveSetting "6", "pw", "3", Text14.Text
SaveSetting "7", "id", "3", Text13.Text
SaveSetting "7", "pw", "3", Text12.Text
SaveSetting "8", "id", "3", Text11.Text
SaveSetting "8", "pw", "3", Text10.Text
SaveSetting "9", "id", "3", Text18.Text
SaveSetting "9", "pw", "3", Text19.Text
SaveSetting "10", "id", "3", Text20.Text
SaveSetting "10", "pw", "3", Text21.Text
SaveSetting "11", "id", "3", Text22.Text
SaveSetting "11", "pw", "3", Text23.Text
SaveSetting "12", "id", "3", Text24.Text
SaveSetting "12", "pw", "3", Text25.Text
SaveSetting "1", "site", "3", Text32.Text
SaveSetting "2", "site", "3", Text33.Text
SaveSetting "3", "site", "3", Text31.Text
SaveSetting "4", "site", "3", Text30.Text
SaveSetting "5", "site", "3", Text34.Text
SaveSetting "6", "site", "3", Text35.Text
SaveSetting "7", "site", "3", Text36.Text
SaveSetting "8", "site", "3", Text37.Text
SaveSetting "9", "site", "3", Text26.Text
SaveSetting "10", "site", "3", Text27.Text
SaveSetting "11", "site", "3", Text28.Text
SaveSetting "12", "site", "3", Text29.Text
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting save Success" & vbCrLf
MsgBox "저장을 성공하였습니다.", vbInformation, "알림"
Else
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting save Fail" & vbCrLf
MsgBox "사용자가 거절하여 저장을 실패하엿습니다", vbInformation, "알림"
End If
End Sub
Private Sub Command14_Click()
Dim a
a = Date + Time
Text9.Text = Text9.Text & "/" & a & "/" & ":ID LIST" & vbCrLf
List.Show
End Sub

Private Sub Command15_Click()
SettingMagician.Show
main.Hide
End Sub

Private Sub Command5_Click()
Frame6.Left = 2280
Frame6.Top = 480
Frame4.Left = 8520
Frame4.Top = 480
Frame8.Left = 8520
Frame8.Top = 480
End Sub

Private Sub Command7_Click()
Frame6.Left = 8520
Frame6.Top = 8520
Frame4.Left = 8520
Frame4.Top = 480
Frame8.Left = 2280
Frame8.Top = 480
End Sub

Private Sub Command8_Click()
Dim a
a = Date + Time
If MsgBox("정말로 저장을 하시겟습니까?", vbYesNo, "알림") = vbYes Then
SaveSetting "1", "id", "3", Text1.Text
SaveSetting "1", "pw", "3", Text2.Text
SaveSetting "2", "id", "3", Text3.Text
SaveSetting "2", "pw", "3", Text4.Text
SaveSetting "3", "id", "3", Text5.Text
SaveSetting "3", "pw", "3", Text6.Text
SaveSetting "4", "id", "3", Text7.Text
SaveSetting "4", "pw", "3", Text8.Text
SaveSetting "5", "id", "3", Text17.Text
SaveSetting "5", "pw", "3", Text16.Text
SaveSetting "6", "id", "3", Text15.Text
SaveSetting "6", "pw", "3", Text14.Text
SaveSetting "7", "id", "3", Text13.Text
SaveSetting "7", "pw", "3", Text12.Text
SaveSetting "8", "id", "3", Text11.Text
SaveSetting "8", "pw", "3", Text10.Text
SaveSetting "9", "id", "3", Text18.Text
SaveSetting "9", "pw", "3", Text19.Text
SaveSetting "10", "id", "3", Text20.Text
SaveSetting "10", "pw", "3", Text21.Text
SaveSetting "11", "id", "3", Text22.Text
SaveSetting "11", "pw", "3", Text23.Text
SaveSetting "12", "id", "3", Text24.Text
SaveSetting "12", "pw", "3", Text25.Text
SaveSetting "1", "site", "3", Text32.Text
SaveSetting "2", "site", "3", Text33.Text
SaveSetting "3", "site", "3", Text31.Text
SaveSetting "4", "site", "3", Text30.Text
SaveSetting "5", "site", "3", Text34.Text
SaveSetting "6", "site", "3", Text35.Text
SaveSetting "7", "site", "3", Text36.Text
SaveSetting "8", "site", "3", Text37.Text
SaveSetting "9", "site", "3", Text26.Text
SaveSetting "10", "site", "3", Text27.Text
SaveSetting "11", "site", "3", Text28.Text
SaveSetting "12", "site", "3", Text29.Text
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting save Success" & vbCrLf
MsgBox "저장을 성공하였습니다.", vbInformation, "알림"
Else
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting save Fail" & vbCrLf
MsgBox "사용자가 거절하여 저장을 실패하엿습니다", vbInformation, "알림"
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If start.Enabled = True Then
Cancel = 1
MsgBox "중지하신후 프로그램을 종료를 하십시오.", vbInformation, "알림"
Else
End
End If
End Sub

Private Sub Label1_Click()

End Sub

Private Sub start_Timer()
setting.Text1.Locked = True
setting.Text2.Locked = True
setting.Text3.Locked = True
setting.Text4.Locked = True
setting.Text11.Locked = True
setting.Text10.Locked = True
setting.Text9.Locked = True
setting.Text8.Locked = True
setting.Text12.Locked = True
setting.Text13.Locked = True
setting.Text14.Locked = True
setting.Text15.Locked = True
setting.Command1.Enabled = False
setting.Command2.Enabled = False
setting.Command9.Enabled = False
setting.Command8.Enabled = False
setting.Command15.Enabled = False
setting.Command16.Enabled = False
End Sub

Private Sub stop1_Timer()
stop1.Enabled = False
start.Enabled = False
End Sub

Private Sub Timer1_Timer()
Dim a
a = Date + Time
If GetAsyncKeyState(setting.Text1.Text) And GetAsyncKeyState(35) Then
If Check1.value = 1 Then
SendKeys (Chr(8))
SendKeys (Text1.Text)
Timer2.Enabled = True
Timer1.Enabled = False
Else
Text9.Text = Text9.Text & "/" & a & "/" & ":Setting - Auto Login Fail" & vbCrLf
End If
End If

End Sub

Private Sub Check1_Click()
Dim a
a = Date + Time
If Check1.value = 1 Then
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting - Setting1 Check Enable" & vbCrLf
SaveSetting "ch", "1", "3", "che1"
Text1.Enabled = True
Text2.Enabled = True
Text1.Text = ""
Text2.Text = ""
Text32.Text = ""
setting.Text1.Text = ""
Else
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting - Setting1 Check Disable" & vbCrLf
DeleteSetting "ch", "1", "3"
Text1.Text = ""
Text2.Text = ""
Text32.Text = ""
Text1.Enabled = False
Text2.Enabled = False
setting.Text1.Text = ""
End If
End Sub

Private Sub Check2_Click()
Dim a
a = Date + Time
If Check2.value = 1 Then
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting - Setting2 Check Enable" & vbCrLf
SaveSetting "ch", "2", "3", "che2"
Text3.Enabled = True
Text4.Enabled = True
Text3.Text = ""
Text4.Text = ""
Text33.Text = ""
setting.Text2.Text = ""
Else
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting - Setting2 Check Disable" & vbCrLf
DeleteSetting "ch", "2", "3"
Text3.Text = ""
Text4.Text = ""
Text33.Text = ""
Text3.Enabled = False
Text4.Enabled = False
setting.Text2.Text = ""
End If
End Sub

Private Sub Check3_Click()
Dim a
a = Date + Time
If Check3.value = 1 Then
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting - Setting3 Check Enable" & vbCrLf
SaveSetting "ch", "3", "3", "che3"
Text5.Enabled = True
Text6.Enabled = True
Text5.Text = ""
Text6.Text = ""
Text31.Text = ""
setting.Text3.Text = ""
Else
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting - Setting3 Check Disable" & vbCrLf
DeleteSetting "ch", "3", "3"
Text5.Text = ""
Text6.Text = ""
Text31.Text = ""
Text5.Enabled = False
Text6.Enabled = False
setting.Text3.Text = ""
End If
End Sub

Private Sub Check4_Click()
Dim a
a = Date + Time
If Check4.value = 1 Then
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting - Setting4 Check Enable" & vbCrLf
SaveSetting "ch", "4", "3", "che4"
Text7.Enabled = True
Text8.Enabled = True
Text7.Text = ""
Text8.Text = ""
Text33.Text = ""
setting.Text4.Text = ""
Else
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting - Setting4 Check Disable" & vbCrLf
DeleteSetting "ch", "4", "3"
Text7.Text = ""
Text8.Text = ""
Text33.Text = ""
Text7.Enabled = False
Text8.Enabled = False
setting.Text4.Text = ""
End If
End Sub

Private Sub Command10_Click()
Dim a
a = Date + Time
If Check1.value = 0 Then
Else
If setting.Text1.Text = "" Then
GoTo w
Else
Timer1.Enabled = True
End If
End If

If Check2.value = 0 Then
Else
If setting.Text2.Text = "" Then
GoTo w
Else
Timer3.Enabled = True
End If
End If

If Check3.value = 0 Then
Else
If setting.Text3.Text = "" Then
GoTo w
Else
Timer5.Enabled = True
End If
End If

If Check4.value = 0 Then
Else
If setting.Text4.Text = "" Then
GoTo w
Else
Timer7.Enabled = True
End If
End If

If Check5.value = 0 Then
Else
If setting.Text11.Text = "" Then
GoTo w
Else
Timer9.Enabled = True
End If
End If

If Check6.value = 0 Then
Else
If setting.Text10.Text = "" Then
GoTo w
Else
Timer11.Enabled = True
End If
End If

If Check7.value = 0 Then
Else
If setting.Text9.Text = "" Then
GoTo w
Else
Timer13.Enabled = True
End If
End If

If Check8.value = 0 Then
Else
If setting.Text8.Text = "" Then
GoTo w
Else
Timer15.Enabled = True
End If
End If

If Check9.value = 0 Then
Else
If setting.Text12.Text = "" Then
GoTo w
Else
Timer17.Enabled = True
End If
End If

If Check10.value = 0 Then
Else
If setting.Text13.Text = "" Then
GoTo w
Else
Timer19.Enabled = True
End If
End If

If Check11.value = 0 Then
Else
If setting.Text14.Text = "" Then
GoTo w
Else
Timer21.Enabled = True
End If
End If

If Check12.value = 0 Then
Else
If setting.Text15.Text = "" Then
GoTo w
Else
Timer23.Enabled = True
End If
End If

Text9.Text = Text9.Text & "/" & a & "/" & ":Automatic Login Program Start...." & vbCrLf
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Text6.Locked = True
Text7.Locked = True
Text8.Locked = True
Text9.Locked = True
Text10.Locked = True
Text11.Locked = True
Text12.Locked = True
Check1.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
Check4.Enabled = False
Check5.Enabled = False
Check6.Enabled = False
Check7.Enabled = False
Check8.Enabled = False
Check9.Enabled = False
Check10.Enabled = False
Check11.Enabled = False
Check12.Enabled = False
Command2.Enabled = False
Command9.Enabled = True
Command10.Enabled = False
setting.Hide
start.Enabled = True
stop1.Enabled = False
GoTo j
w:
MsgBox "핫키가 지정되지 않았습니다"
Text9.Text = Text9.Text & "/" & a & "/" & ":Starting fail - Hotkey setting" & vbCrLf
j:
End Sub

Private Sub Command3_Click()
Update.Show
Dim a
a = Date + Time
Text9.Text = Text9.Text & "/" & a & "/" & ":Update Check" & vbCrLf
End Sub



Private Sub Command1_Click()
Dim a
a = Date + Time
Text9.Text = Text9.Text & "/" & a & "/" & ":Program information" & vbCrLf
programinfo.Show
End Sub

Private Sub Command2_Click()
Dim a
a = Date + Time
Text9.Text = Text9.Text & "/" & a & "/" & ":Setting" & vbCrLf
If Command2.Enabled = True Then
setting.Show
Else
MsgBox "시작중에는 설정을 할수가없습니다.중지를하신후 설정을 해주세요", vbInformation, "안내"
End If
End Sub


Private Sub Command4_Click()
Dim a
a = Date + Time
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting" & vbCrLf
Frame2.Left = 8520
Frame4.Left = 2280
Frame6.Left = 8520
Frame6.Top = 480
End Sub

Private Sub Command6_Click()
Dim a
a = Date + Time
If MsgBox("정말로 저장을 하시겟습니까?", vbYesNo, "알림") = vbYes Then
SaveSetting "1", "id", "3", Text1.Text
SaveSetting "1", "pw", "3", Text2.Text
SaveSetting "2", "id", "3", Text3.Text
SaveSetting "2", "pw", "3", Text4.Text
SaveSetting "3", "id", "3", Text5.Text
SaveSetting "3", "pw", "3", Text6.Text
SaveSetting "4", "id", "3", Text7.Text
SaveSetting "4", "pw", "3", Text8.Text
SaveSetting "5", "id", "3", Text17.Text
SaveSetting "5", "pw", "3", Text16.Text
SaveSetting "6", "id", "3", Text15.Text
SaveSetting "6", "pw", "3", Text14.Text
SaveSetting "7", "id", "3", Text13.Text
SaveSetting "7", "pw", "3", Text12.Text
SaveSetting "8", "id", "3", Text11.Text
SaveSetting "8", "pw", "3", Text10.Text
SaveSetting "9", "id", "3", Text18.Text
SaveSetting "9", "pw", "3", Text19.Text
SaveSetting "10", "id", "3", Text20.Text
SaveSetting "10", "pw", "3", Text21.Text
SaveSetting "11", "id", "3", Text22.Text
SaveSetting "11", "pw", "3", Text23.Text
SaveSetting "12", "id", "3", Text24.Text
SaveSetting "12", "pw", "3", Text25.Text
SaveSetting "1", "site", "3", Text32.Text
SaveSetting "2", "site", "3", Text33.Text
SaveSetting "3", "site", "3", Text31.Text
SaveSetting "4", "site", "3", Text30.Text
SaveSetting "5", "site", "3", Text34.Text
SaveSetting "6", "site", "3", Text35.Text
SaveSetting "7", "site", "3", Text36.Text
SaveSetting "8", "site", "3", Text37.Text
SaveSetting "9", "site", "3", Text26.Text
SaveSetting "10", "site", "3", Text27.Text
SaveSetting "11", "site", "3", Text28.Text
SaveSetting "12", "site", "3", Text29.Text
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting save Success" & vbCrLf
MsgBox "저장을 성공하였습니다.", vbInformation, "알림"
Else
Text9.Text = Text9.Text & "/" & a & "/" & ":ID Setting save Fail" & vbCrLf
MsgBox "사용자가 거절하여 저장을 실패하엿습니다", vbInformation, "알림"
End If

End Sub

Private Sub Command9_Click()
Timer1.Enabled = False
Dim a
a = Date + Time
Text9.Text = Text9.Text & "/" & a & "/" & ":Automatic Login Program Stop...." & vbCrLf
start.Enabled = False
stop1.Enabled = True
setting.Command1.Enabled = True
setting.Command2.Enabled = True
setting.Command9.Enabled = True
setting.Command8.Enabled = True
setting.Command15.Enabled = True
setting.Command16.Enabled = True
Command9.Enabled = False
Command10.Enabled = True
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Text5.Locked = False
Text6.Locked = False
Text7.Locked = False
Text8.Locked = False
Text9.Locked = False
Text10.Locked = False
Text11.Locked = False
Text12.Locked = False
Check1.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Check4.Enabled = True
Check5.Enabled = True
Check6.Enabled = True
Check7.Enabled = True
Check8.Enabled = True
Check9.Enabled = True
Check10.Enabled = True
Check11.Enabled = True
Check12.Enabled = True
Command2.Enabled = True
Command15.Enabled = True
End Sub

Private Sub Form_Load()
setting.Show
setting.Hide
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
Text12.Enabled = False
Text13.Enabled = False
Text14.Enabled = False
Text15.Enabled = False
Text16.Enabled = False
Text17.Enabled = False
Text18.Enabled = False
Text19.Enabled = False
Text20.Enabled = False
Text21.Enabled = False
Text22.Enabled = False
Text23.Enabled = False
Text24.Enabled = False
Text25.Enabled = False
If GetSetting("ch", "1", "3") = "che1" Then
Check1.value = 1
Else
Check1.value = 0
End If
If GetSetting("ch", "2", "3") = "che2" Then
Check2.value = 1
Else
Check2.value = 0
End If
If GetSetting("ch", "3", "3") = "che3" Then
Check3.value = 1
Else
Check3.value = 0
End If
If GetSetting("ch", "4", "3") = "che4" Then
Check4.value = 1
Else
Check4.value = 0
End If
If GetSetting("ch", "5", "3") = "che5" Then
Check5.value = 1
Else
Check5.value = 0
End If
If GetSetting("ch", "6", "3") = "che6" Then
Check6.value = 1
Else
Check6.value = 0
End If
If GetSetting("ch", "7", "3") = "che7" Then
Check7.value = 1
Else
Check7.value = 0
End If
If GetSetting("ch", "8", "3") = "che8" Then
Check8.value = 1
Else
Check8.value = 0
End If
If GetSetting("ch", "9", "3") = "che9" Then
Check9.value = 1
Else
Check9.value = 0
End If
If GetSetting("ch", "10", "3") = "che10" Then
Check10.value = 1
Else
Check10.value = 0
End If
If GetSetting("ch", "11", "3") = "che11" Then
Check11.value = 1
Else
Check11.value = 0
End If
If GetSetting("ch", "12", "3") = "che12" Then
Check12.value = 1
Else
Check12.value = 0
End If

If Check1.value = 1 Then
Text1.Text = GetSetting("1", "id", "3")
Text2.Text = GetSetting("1", "pw", "3")
Text32.Text = GetSetting("1", "site", "3")
End If
If Check2.value = 1 Then
Text3.Text = GetSetting("2", "id", "3")
Text4.Text = GetSetting("2", "pw", "3")
Text33.Text = GetSetting("2", "site", "3")
End If
If Check3.value = 1 Then
Text5.Text = GetSetting("3", "id", "3")
Text6.Text = GetSetting("3", "pw", "3")
Text31.Text = GetSetting("3", "site", "3")
End If
If Check4.value = 1 Then
Text7.Text = GetSetting("4", "id", "3")
Text8.Text = GetSetting("4", "pw", "3")
Text30.Text = GetSetting("4", "site", "3")
End If
If Check5.value = 1 Then
Text17.Text = GetSetting("5", "id", "3")
Text16.Text = GetSetting("5", "pw", "3")
Text34.Text = GetSetting("5", "site", "3")
End If
If Check6.value = 1 Then
Text15.Text = GetSetting("6", "id", "3")
Text14.Text = GetSetting("6", "pw", "3")
Text35.Text = GetSetting("6", "site", "3")
End If
If Check7.value = 1 Then
Text13.Text = GetSetting("7", "id", "3")
Text12.Text = GetSetting("7", "pw", "3")
Text36.Text = GetSetting("7", "site", "3")
End If
If Check8.value = 1 Then
Text11.Text = GetSetting("8", "id", "3")
Text10.Text = GetSetting("8", "pw", "3")
Text37.Text = GetSetting("8", "site", "3")
End If
If Check9.value = 1 Then
Text18.Text = GetSetting("9", "id", "3")
Text19.Text = GetSetting("9", "pw", "3")
Text26.Text = GetSetting("9", "site", "3")
End If
If Check10.value = 1 Then
Text20.Text = GetSetting("10", "id", "3")
Text21.Text = GetSetting("10", "pw", "3")
Text27.Text = GetSetting("10", "site", "3")
End If
If Check11.value = 1 Then
Text22.Text = GetSetting("11", "id", "3")
Text23.Text = GetSetting("11", "pw", "3")
Text28.Text = GetSetting("11", "site", "3")
End If
If Check12.value = 1 Then
Text24.Text = GetSetting("12", "id", "3")
Text25.Text = GetSetting("12", "pw", "3")
Text29.Text = GetSetting("12", "site", "3")
End If
If GetSetting("a", "u", "t") = "az" Then
Command10.value = True
End If
End Sub
Private Sub Timer10_Timer()
Dim a
a = Date + Time
SendKeys (Chr(vbKeyTab))
SendKeys (Text16.Text)
SendKeys (Chr(13))
Text9.Text = Text9.Text & "/" & a & "/" & ":Setting - Auto Login Success" & vbCrLf
Timer10.Enabled = False
Timer9.Enabled = True
End Sub

Private Sub Timer11_Timer()
Dim a
a = Date + Time
If GetAsyncKeyState(setting.Text10.Text) And GetAsyncKeyState(17) Then
If Check6.value = 1 Then
SendKeys (Text15.Text)
Timer12.Enabled = True
Timer11.Enabled = False
Else
Text9.Text = Text9.Text & "/" & a & "/" & ":Setting - Auto Login Fail" & vbCrLf
End If
End If
End Sub

Private Sub Timer12_Timer()
Dim a
a = Date + Time
SendKeys (Chr(vbKeyTab))
SendKeys (Text14.Text)
SendKeys (Chr(13))
Text9.Text = Text9.Text & "/" & a & "/" & ":Setting - Auto Login Success" & vbCrLf
Timer12.Enabled = False
Timer11.Enabled = True

End Sub

Private Sub Timer13_Timer()

Dim a
a = Date + Time
If GetAsyncKeyState(setting.Text9.Text) And GetAsyncKeyState(17) Then
If Check7.value = 1 Then
SendKeys (Text13.Text)
Timer14.Enabled = True
Timer13.Enabled = False
Else
Text9.Text = Text9.Text & "/" & a & "/" & ":Setting - Auto Login Fail" & vbCrLf
End If
End If
End Sub

Private Sub Timer14_Timer()
Dim a
a = Date + Time
SendKeys (Chr(vbKeyTab))
SendKeys (Text12.Text)
SendKeys (Chr(13))
Text9.Text = Text9.Text & "/" & a & "/" & ":Setting - Auto Login Success" & vbCrLf
Timer14.Enabled = False
Timer13.Enabled = True

End Sub

Private Sub Timer15_Timer()

Dim a
a = Date + Time
If GetAsyncKeyState(setting.Text8.Text) And GetAsyncKeyState(17) Then
If Check8.value = 1 Then
SendKeys (Text11.Text)
Timer16.Enabled = True
Timer15.Enabled = False
Else
Text9.Text = Text9.Text & "/" & a & "/" & ":Setting - Auto Login Fail" & vbCrLf
End If
End If
End Sub

Private Sub Timer16_Timer()
Dim a
a = Date + Time
SendKeys (Chr(vbKeyTab))
SendKeys (Text10.Text)
SendKeys (Chr(13))
Text9.Text = Text9.Text & "/" & a & "/" & ":Setting - Auto Login Success" & vbCrLf
Timer16.Enabled = False
Timer15.Enabled = True

End Sub

Private Sub Timer17_Timer()
Dim a
a = Date + Time
If GetAsyncKeyState(setting.Text12.Text) And GetAsyncKeyState(17) Then
If Check9.value = 1 Then
SendKeys (Text18.Text)
Timer18.Enabled = True
Timer17.Enabled = False
Else
Text9.Text = Text9.Text & "/" & a & "/" & ":Setting - Auto Login Fail" & vbCrLf
End If
End If
End Sub

Private Sub Timer18_Timer()
Dim a
a = Date + Time
SendKeys (Chr(vbKeyTab))
SendKeys (Text19.Text)
SendKeys (Chr(13))
Text9.Text = Text9.Text & "/" & a & "/" & ":Setting - Auto Login Success" & vbCrLf
Timer18.Enabled = False
Timer17.Enabled = True

End Sub

Private Sub Timer19_Timer()
Dim a
a = Date + Time
If GetAsyncKeyState(setting.Text13.Text) And GetAsyncKeyState(17) Then
If Check10.value = 1 Then
SendKeys (Text20.Text)
Timer20.Enabled = True
Timer19.Enabled = False
Else
Text9.Text = Text9.Text & "/" & a & "/" & ":Setting - Auto Login Fail" & vbCrLf
End If
End If
End Sub

Private Sub Timer2_Timer()
Dim a
a = Date + Time
SendKeys (Chr(vbKeyTab))
SendKeys (Text2.Text)
SendKeys (Chr(13))
Text9.Text = Text9.Text & "/" & a & "/" & ":Setting - Auto Login Success" & vbCrLf
Timer2.Enabled = False
Timer1.Enabled = True
End Sub

Private Sub Timer20_Timer()
Dim a
a = Date + Time
SendKeys (Chr(vbKeyTab))
SendKeys (Text21.Text)
SendKeys (Chr(13))
Text9.Text = Text9.Text & "/" & a & "/" & ":Setting - Auto Login Success" & vbCrLf
Timer20.Enabled = False
Timer19.Enabled = True

End Sub

Private Sub Timer21_Timer()
Dim a
a = Date + Time
If GetAsyncKeyState(setting.Text14.Text) And GetAsyncKeyState(17) Then
If Check11.value = 1 Then
SendKeys (Text22.Text)
Timer22.Enabled = True
Timer21.Enabled = False
Else
Text9.Text = Text9.Text & "/" & a & "/" & ":Setting - Auto Login Fail" & vbCrLf
End If
End If
End Sub

Private Sub Timer22_Timer()
Dim a
a = Date + Time
SendKeys (Chr(vbKeyTab))
SendKeys (Text23.Text)
SendKeys (Chr(13))
Text9.Text = Text9.Text & "/" & a & "/" & ":Setting - Auto Login Success" & vbCrLf
Timer22.Enabled = False
Timer21.Enabled = True


End Sub

Private Sub Timer23_Timer()

Dim a
a = Date + Time
If GetAsyncKeyState(setting.Text15.Text) And GetAsyncKeyState(17) Then
If Check12.value = 1 Then
SendKeys (Text24.Text)
Timer24.Enabled = True
Timer23.Enabled = False
Else
Text9.Text = Text9.Text & "/" & a & "/" & ":Setting - Auto Login Fail" & vbCrLf
End If
End If
End Sub

Private Sub Timer24_Timer()
Dim a
a = Date + Time
SendKeys (Chr(vbKeyTab))
SendKeys (Text25.Text)
SendKeys (Chr(13))
Text9.Text = Text9.Text & "/" & a & "/" & ":Setting - Auto Login Success" & vbCrLf
Timer24.Enabled = False
Timer23.Enabled = True
End Sub

Private Sub Timer25_Timer()
Dim a
a = Date + Time
If GetAsyncKeyState(16) And GetAsyncKeyState(9) Then
If List.Visible = True Then
Else
List.Show
Text9.Text = Text9.Text & "/" & a & "/" & ":ID List" & vbCrLf
End If
End If
End Sub

Private Sub Timer3_Timer()
Dim a
a = Date + Time
If GetAsyncKeyState(setting.Text2.Text) And GetAsyncKeyState(17) Then
If Check2.value = 1 Then
SendKeys (Text3.Text)
Timer4.Enabled = True
Timer3.Enabled = False
Else
Text9.Text = Text9.Text & "/" & a & "/" & ":Setting - Auto Login Fail" & vbCrLf
End If
End If

End Sub

Private Sub Timer4_Timer()
Dim a
a = Date + Time
SendKeys (Chr(vbKeyTab))
SendKeys (Text4.Text)
SendKeys (Chr(13))
Text9.Text = Text9.Text & "/" & a & "/" & ":Setting - Auto Login Success" & vbCrLf
Timer4.Enabled = False
Timer3.Enabled = True
End Sub

Private Sub Timer5_Timer()
Dim a
a = Date + Time
If GetAsyncKeyState(setting.Text3.Text) And GetAsyncKeyState(17) Then
If Check3.value = 1 Then
SendKeys (Text5.Text)
Timer6.Enabled = True
Timer5.Enabled = False
Else
Text9.Text = Text9.Text & "/" & a & "/" & ":Setting - Auto Login Fail" & vbCrLf
End If
End If

End Sub

Private Sub Timer6_Timer()
Dim a
a = Date + Time
SendKeys (Chr(vbKeyTab))
SendKeys (Text6.Text)
SendKeys (Chr(13))
Text9.Text = Text9.Text & "/" & a & "/" & ":Setting - Auto Login Success" & vbCrLf
Timer6.Enabled = False
Timer5.Enabled = True
End Sub

Private Sub Timer7_Timer()
Dim a
a = Date + Time
If GetAsyncKeyState(setting.Text4.Text) And GetAsyncKeyState(17) Then
If Check4.value = 1 Then
SendKeys (Text7.Text)
Timer8.Enabled = True
Timer7.Enabled = False
Else
Text9.Text = Text9.Text & "/" & a & "/" & ":Setting - Auto Login Fail" & vbCrLf
End If
End If
End Sub

Private Sub Timer8_Timer()
Dim a
a = Date + Time
SendKeys (Chr(vbKeyTab))
SendKeys (Text8.Text)
SendKeys (Chr(13))
Text9.Text = Text9.Text & "/" & a & "/" & ":Setting - Auto Login Success" & vbCrLf
Timer8.Enabled = False
Timer7.Enabled = True
End Sub

Private Sub Timer9_Timer()
Dim a
a = Date + Time
If GetAsyncKeyState(setting.Text11.Text) And GetAsyncKeyState(17) Then
If Check5.value = 1 Then
SendKeys (Text17.Text)
Timer10.Enabled = True
Timer9.Enabled = False
Else
Text9.Text = Text9.Text & "/" & a & "/" & ":Setting - Auto Login Fail" & vbCrLf
End If
End If
End Sub
