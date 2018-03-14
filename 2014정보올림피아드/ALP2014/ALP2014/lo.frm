VERSION 5.00
Begin VB.Form lo 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   3900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin Project1.ctlSkin ctlSkin1 
      Height          =   4335
      Left            =   -120
      TabIndex        =   0
      Top             =   -360
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   7646
      Caption         =   "lo"
      Resize          =   0   'False
      Begin VB.Timer Timer1 
         Interval        =   2000
         Left            =   2040
         Top             =   3000
      End
      Begin VB.Image Image2 
         Height          =   435
         Left            =   7320
         Picture         =   "lo.frx":0000
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   1125
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "Ver:1.0.0"
         BeginProperty Font 
            Name            =   "a피오피네모"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   3
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "(ALP2014)"
         BeginProperty Font 
            Name            =   "a피오피네모"
            Size            =   15.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4800
         TabIndex        =   2
         Top             =   2040
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "Automatic Login Program"
         BeginProperty Font 
            Name            =   "a피오피네모"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3000
         TabIndex        =   1
         Top             =   1560
         Width           =   6375
      End
      Begin VB.Image Image1 
         Height          =   1920
         Left            =   360
         Picture         =   "lo.frx":C325
         Stretch         =   -1  'True
         Top             =   960
         Width           =   2160
      End
   End
End
Attribute VB_Name = "lo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
agree.Show
lo.Hide
End Sub
