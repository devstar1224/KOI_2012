VERSION 5.00
Begin VB.UserControl ctlSkin 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   2010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3075
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000007&
   ScaleHeight     =   134
   ScaleMode       =   3  '픽셀
   ScaleWidth      =   205
   Begin VB.PictureBox picButtons 
      Appearance      =   0  '평면
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   1200
      Picture         =   "ctlSkin.ctx":0000
      ScaleHeight     =   145
      ScaleMode       =   0  '사용자
      ScaleWidth      =   102
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.PictureBox picLayout 
      Appearance      =   0  '평면
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   120
      Picture         =   "ctlSkin.ctx":AEB8
      ScaleHeight     =   64
      ScaleMode       =   0  '사용자
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "ctlSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const cWidth = 1
Private Const uHeight = 1
Private Const dHeight = 1
Const bHeight = 29
Const bMargin = 0
Const bMinWidth = 32
Const bMaxWidth = 34
Const bExitWidth = 33
Const cx = 9
Const cy = 4
Private bLeft As Integer
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const SC_DRAG_RESIZEL As Long = 10&
Private Const SC_DRAG_RESIZER As Long = 11&
Private Const SC_DRAG_RESIZEU As Long = 12&
Private Const SC_DRAG_RESIZEUL As Long = 13&
Private Const SC_DRAG_RESIZEUR As Long = 14&
Private Const SC_DRAG_RESIZED As Long = 15&
Private Const SC_DRAG_RESIZEDL As Long = 16&
Private Const SC_DRAG_RESIZEDR As Long = 17&
Private Const SC_DRAG_MOVE As Long = 2&
Dim cName As String
Dim bResize As Boolean



Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Dim wArea As RECT



Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type PointAPI
    X As Long
    Y As Long
End Type

Private Type BtnPosition
    X As Single
    Y As Single
End Type


Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function TextOut Lib "gdi32.dll" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long


Dim bMaxed As Boolean
Dim originalRect As RECT



Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As PointAPI) As Long
Private Declare Function SetCapture Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetCapture Lib "user32.dll" () As Long
Dim iPA As PointAPI, iRect As RECT, iBP As BtnPosition

Dim bButtons As Boolean
Dim bMinimize As Boolean, bMaximize As Boolean, bExit As Boolean
Dim g_MouseIn As Boolean, g_MouseDown As Boolean
Dim iMouseDown As Integer

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Sub drawForm()
    On Error Resume Next
    Cls
    BitBlt hdc, 0, 0, cWidth, uHeight, picLayout.hdc, 0, 0, vbSrcCopy
    StretchBlt hdc, cWidth, 0, Width / 15 - 2 * cWidth, uHeight, picLayout.hdc, cWidth, 0, picLayout.Width - 2 * cWidth, uHeight, vbSrcCopy
    BitBlt hdc, Width / 15 - cWidth, 0, cWidth, uHeight, picLayout.hdc, picLayout.Width - cWidth, 0, vbSrcCopy
    
    StretchBlt hdc, 0, uHeight, cWidth, Height / 15 - uHeight - dHeight, picLayout.hdc, 0, uHeight, cWidth, picLayout.Height - uHeight - dHeight, vbSrcCopy
    StretchBlt hdc, cWidth, uHeight, Width / 15 - 20, Height / 15 - uHeight - dHeight, picLayout.hdc, cWidth, uHeight, picLayout.Width - 2 * cWidth, picLayout.Height - uHeight - dHeight, vbSrcCopy
    StretchBlt hdc, Width / 15 - cWidth, uHeight, cWidth, Height / 15 - uHeight - dHeight, picLayout.hdc, picLayout.Width - cWidth, uHeight, cWidth, picLayout.Height - uHeight - dHeight, vbSrcCopy
    
    BitBlt hdc, 0, Height / 15 - dHeight, cWidth, dHeight, picLayout.hdc, 0, picLayout.Height - dHeight, vbSrcCopy
    StretchBlt hdc, cWidth, Height / 15 - dHeight, Width / 15 - 2 * cWidth, dHeight, picLayout.hdc, cWidth, picLayout.Height - dHeight, cWidth, dHeight, vbSrcCopy
    BitBlt hdc, Width / 15 - cWidth, Height / 15 - dHeight, cWidth, dHeight, picLayout.hdc, picLayout.Width - cWidth, picLayout.Height - dHeight, vbSrcCopy


    BitBlt hdc, Width / 15 - bLeft, 0, bMinWidth, bHeight, picButtons.hdc, 0, 0, vbSrcCopy
    If bResize = True Then
        If bMaxed = True Then
            BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, 0, bHeight * 4, vbSrcCopy
        Else
            BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, bMinWidth, 0, vbSrcCopy
        End If
    Else
        BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, bMinWidth, bHeight * 3, vbSrcCopy
    End If
    BitBlt hdc, Width / 15 - bLeft + bMinWidth + bMaxWidth, 0, bExitWidth, bHeight, picButtons.hdc, bMinWidth + bMaxWidth, 0, vbSrcCopy

        
        TextOut hdc, cx, cy, cName, LenB(StrConv(cName, vbFromUnicode))
    
End Sub

Private Sub UserControl_DblClick()
If Resize = False Then Exit Sub

If (iRect.Right - iPA.X > bLeft And iPA.Y - iRect.Top < uHeight) Then
        If bMaxed = False Then

            originalRect.Left = Parent.Left
            originalRect.Right = Parent.Width
            originalRect.Top = Parent.Top
            originalRect.Bottom = Parent.Height
            
            
            SystemParametersInfo 48, 0, wArea, 0
            
            bMaxed = True
            Parent.Left = wArea.Left * 15
            Parent.Top = wArea.Top * 15
            Parent.Width = (wArea.Right - wArea.Left) * 15
            Parent.Height = (wArea.Bottom - wArea.Top) * 15
            Width = Parent.Width
            Height = Parent.Height
            drawForm
            Exit Sub
        Else
            Parent.Left = originalRect.Left
            Parent.Width = originalRect.Right
            Parent.Top = originalRect.Top
            Parent.Height = originalRect.Bottom
            Width = Parent.Width
            Height = Parent.Height
            bMaxed = False
            drawForm
            
        End If
End If
End Sub

Sub drawButtons()
GetCursorPos iPA
GetWindowRect hwnd, iRect



Dim a, B, C, d
a = iPA.X - iRect.Left
B = iPA.Y - iRect.Top
C = iRect.Right - iPA.X
d = iRect.Bottom - iPA.Y

Dim intRight As Integer
intRight = bLeft - bMinWidth - bMaxWidth - bExitWidth





If (0 < B And B < bHeight) And (intRight < C And C < bLeft) Then
    bButtons = True
    
    If C - intRight <= bExitWidth Then
        bExit = True
        bMinimize = False
        bMaximize = False
        
        If iMouseDown = 1 Then
           BitBlt hdc, Width / 15 - bLeft + bMinWidth + bMaxWidth, 0, bExitWidth, bHeight, picButtons.hdc, bMinWidth + bMaxWidth, 2 * bHeight, vbSrcCopy
        Else
           
            BitBlt hdc, Width / 15 - bLeft + bMinWidth + bMaxWidth, 0, bExitWidth, bHeight, picButtons.hdc, bMinWidth + bMaxWidth, bHeight, vbSrcCopy
        End If
        
      
        If bResize = True Then
            If bMaxed = True Then
                BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, 0, bHeight * 4, vbSrcCopy
            Else
                BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, bMinWidth, 0, vbSrcCopy
            End If
        Else
            BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, bMinWidth, bHeight * 3, vbSrcCopy
        End If
        BitBlt hdc, Width / 15 - bLeft, 0, bMinWidth, bHeight, picButtons.hdc, 0, 0, vbSrcCopy
        
    
    ElseIf C - intRight <= bExitWidth + bMaxWidth Then
        bMaximize = True
        bMinimize = False
        bExit = False
        
        
   
        If bResize = True Then
            If iMouseDown = 1 Then
                If bMaxed = True Then
                     BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, bMaxWidth, bHeight * 4, vbSrcCopy
                Else
                     BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, bMinWidth, bHeight * 2, vbSrcCopy
                End If
            Else
                If bMaxed = True Then
                     BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, bMaxWidth * 2, bHeight * 4, vbSrcCopy
                Else
                     BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, bMinWidth, bHeight, vbSrcCopy
                End If
                
            End If
        End If
        
 
        BitBlt hdc, Width / 15 - bLeft + bMinWidth + bMaxWidth, 0, bExitWidth, bHeight, picButtons.hdc, bMinWidth + bMaxWidth, 0, vbSrcCopy
        BitBlt hdc, Width / 15 - bLeft, 0, bMinWidth, bHeight, picButtons.hdc, 0, 0, vbSrcCopy
    ElseIf C - intRight <= bExitWidth + bMaxWidth + bMinWidth Then
        bMinimize = True
        bMaximize = False
        bExit = False
        
        

        If iMouseDown = 1 Then
            BitBlt hdc, Width / 15 - bLeft, 0, bMinWidth, bHeight, picButtons.hdc, 0, bHeight * 2, vbSrcCopy
        Else
            BitBlt hdc, Width / 15 - bLeft, 0, bMinWidth, bHeight, picButtons.hdc, 0, bHeight, vbSrcCopy
        End If
        
  
        If bResize = True Then
            If bMaxed = True Then
                BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, 0, bHeight * 4, vbSrcCopy
            Else
                BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, bMinWidth, 0, vbSrcCopy
            End If
        Else
            BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, bMinWidth, bHeight * 3, vbSrcCopy
        End If
        BitBlt hdc, Width / 15 - bLeft + bMinWidth + bMaxWidth, 0, bExitWidth, bHeight, picButtons.hdc, bMinWidth + bMaxWidth, 0, vbSrcCopy
    End If
    
Else
    bButtons = False
    BitBlt hdc, Width / 15 - bLeft, 0, bMinWidth, bHeight, picButtons.hdc, 0, 0, vbSrcCopy
    If bResize = True Then
        If bMaxed = True Then
            BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, 0, bHeight * 4, vbSrcCopy
        Else
            BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, bMinWidth, 0, vbSrcCopy
        End If
    Else
        BitBlt hdc, Width / 15 - bLeft + bMinWidth, 0, bMaxWidth, bHeight, picButtons.hdc, bMinWidth, bHeight * 3, vbSrcCopy
    End If
    BitBlt hdc, Width / 15 - bLeft + bMinWidth + bMaxWidth, 0, bExitWidth, bHeight, picButtons.hdc, bMinWidth + bMaxWidth, 0, vbSrcCopy
    bMinimize = False
    bMaximize = False
    bExit = False
End If
End Sub
Private Sub UserControl_Initialize()
            SystemParametersInfo 48, 0, wArea, 0
            bLeft = bMinWidth + bMaxWidth + bExitWidth + bMargin
End Sub



Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

iMouseDown = 1



    
    Dim SysCommWparam As Integer
    If Button = 1 And bResize = True And bMaxed = False Then
        If (X < 3&) And (Y < 3&) Then
            SysCommWparam = SC_DRAG_RESIZEUL

        ElseIf (X > ScaleWidth - 3&) And (Y > ScaleHeight - 3&) Then
            SysCommWparam = SC_DRAG_RESIZEDR

        ElseIf (X < 3&) And (Y > ScaleHeight - 3&) Then
            SysCommWparam = SC_DRAG_RESIZEDL

        ElseIf (X > ScaleWidth - 3&) And (Y < 3&) Then
            SysCommWparam = SC_DRAG_RESIZEUR

        ElseIf (X < 3&) Then
            SysCommWparam = SC_DRAG_RESIZEL

        ElseIf (X > ScaleWidth - 3&) Then
            SysCommWparam = SC_DRAG_RESIZER

        ElseIf (Y < 3&) Then
            SysCommWparam = SC_DRAG_RESIZEU

        ElseIf (Y > ScaleHeight - 3&) Then
            SysCommWparam = SC_DRAG_RESIZED

        Else
            
        End If

        ReleaseCapture
        SendMessage Parent.hwnd, &HA1, SysCommWparam, 0&
    End If
    

    If Button = 1 And (Y < 30) And bButtons = False And bMaxed = False Then
        SysCommWparam = SC_DRAG_MOVE
        Call ReleaseCapture
        SendMessage Parent.hwnd, &HA1, SysCommWparam, 0&
        Exit Sub
    End If
    
    g_MouseDown = True
    g_MouseIn = True
    
        
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

If Button = 0 And iMouseDown = 1 Then iMouseDown = 0



    If bResize = True And bMaxed = False Then
        If ((X < 3&) And (Y < 3&)) Or ((X > ScaleWidth - 3&) And (Y > ScaleHeight - 3&)) Then
            MousePointer = 8&
    
        ElseIf ((X < 3&) And (Y > ScaleHeight - 3&)) Or ((X > ScaleWidth - 3&) And (Y < 3&)) Then
            MousePointer = 6&
    
        ElseIf ((X < 3&) Or (X > ScaleWidth - 3&)) Then
            MousePointer = 9&
    
        ElseIf ((Y < 3&) Or (Y > ScaleHeight - 3&)) Then
            MousePointer = 7&
    
        Else
            MousePointer = 0&
        End If
    End If
    
    




If Y < 16 Then
    SetCapture hwnd
Else
    ReleaseCapture
End If




If X >= 0 And Y >= 0 And X <= Parent.Width / 15 And Y <= Parent.Height / 15 Then

        
          drawButtons
          
Else
    ReleaseCapture
    drawButtons
    
End If

End Sub







Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If bMinimize = True Then
    Parent.WindowState = 1
    
End If

If bResize = True And bMaximize = True Then
        If bMaxed = False Then
            originalRect.Left = Parent.Left
            originalRect.Right = Parent.Width
            originalRect.Top = Parent.Top
            originalRect.Bottom = Parent.Height
            
            
            SystemParametersInfo 48, 0, wArea, 0
            
            bMaxed = True
            Parent.Left = wArea.Left * 15
            Parent.Top = wArea.Top * 15
            Parent.Width = (wArea.Right - wArea.Left) * 15
            Parent.Height = (wArea.Bottom - wArea.Top) * 15
            Width = Parent.Width
            Height = Parent.Height
            drawForm
        Else
  
            bMaxed = False
            Parent.Left = originalRect.Left
            Parent.Width = originalRect.Right
            Parent.Top = originalRect.Top
            Parent.Height = originalRect.Bottom
            Width = Parent.Width
            Height = Parent.Height
            drawForm
        End If
End If

If bExit = True Then
   Unload Parent
End If
End Sub

Private Sub UserControl_Paint(): drawForm
End Sub

Private Sub UserControl_Resize()
drawForm
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    cName = PropBag.ReadProperty("Caption", UserControl.Name)
    bResize = PropBag.ReadProperty("Resize", True)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Caption", cName, Empty)
Call PropBag.WriteProperty("Resize", bResize, True)
End Sub

Public Property Get Caption() As String
Caption = cName
End Property

Public Property Let Caption(Str As String)
cName = Str
drawForm
End Property

Public Property Get Resize() As Boolean:
Resize = bResize
End Property

Public Property Let Resize(value As Boolean)
bResize = value
PropertyChanged "Resize"
drawForm

End Property



