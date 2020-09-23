VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      DrawWidth       =   10
      Height          =   8175
      Left            =   0
      ScaleHeight     =   8175
      ScaleWidth      =   8055
      TabIndex        =   2
      Top             =   600
      Width           =   8055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   8055
   End
   Begin VB.Label lblmain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Paint Threw"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Lastx As Long
Dim Lasty As Long
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Const LWA_COLORKEY = &H1
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
Const BM_SETSTATE = &HF3

Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Dim ButtonDown As Boolean

Private Sub Command1_Click()
Picture1.Cls
End Sub

Private Sub Form_Load()
Clip
End Sub

Private Sub Label1_Click()
End
End Sub

Private Sub lblmain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Lastx = X
      Lasty = Y
End Sub

Private Sub lblmain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If Button = 1 Then
        Me.Left = Me.Left + (X - Lastx)
        Me.Top = Me.Top + (Y - Lasty)
      End If
End Sub


Sub Clip()
  Dim Ret As Long
  Dim CLR As Long
  CLR = RGB(255, 0, 0) 'this color is the color that will be transparent
  'Set the window style to 'Layered'
  Ret = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
  Ret = Ret Or WS_EX_LAYERED
  SetWindowLong Me.hWnd, GWL_EXSTYLE, Ret
  'Set the opacity of the layered window to 128
  SetLayeredWindowAttributes Me.hWnd, CLR, 0, LWA_COLORKEY
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ButtonDown = True
Picture1.PSet (X, Y), vbRed
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ButtonDown = True Then
Picture1.Line -(X, Y), vbRed
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ButtonDown = False
End Sub
