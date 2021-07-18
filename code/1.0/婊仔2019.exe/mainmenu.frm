VERSION 5.00
Begin VB.Form mainmenu 
   BorderStyle     =   0  'None
   Caption         =   "主菜单"
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleMode       =   0  'User
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "设置"
      Height          =   345
      Left            =   5040
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "新游戏"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "退出游戏"
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   4080
      Width           =   975
   End
End
Attribute VB_Name = "mainmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ZiShiYing_T(), ZiShiYing_L(), ZiShiYing_H(), ZiShiYing_W(), ZiShiYing_N%
Public BGMop As String
Public USER As String

Private Sub Command2_Click()
Unload Me
gmsrt.Show
End Sub

Private Sub Command3_Click()
setting.Show
End Sub

Private Sub Form_Load()
Dim mcontrol As Object
On Error Resume Next
ZiShiYing_N = 0
For Each mcontrol In Me.Controls
If Not (TypeOf mcontrol Is Menu) Then
ZiShiYing_N = ZiShiYing_N + 1
ReDim Preserve ZiShiYing_T(ZiShiYing_N)
ReDim Preserve ZiShiYing_L(ZiShiYing_N)
ReDim Preserve ZiShiYing_H(ZiShiYing_N)
ReDim Preserve ZiShiYing_W(ZiShiYing_N)
ZiShiYing_T(ZiShiYing_N) = mcontrol.Top / ScaleHeight
ZiShiYing_L(ZiShiYing_N) = mcontrol.Left / ScaleWidth
ZiShiYing_H(ZiShiYing_N) = mcontrol.Height / ScaleHeight
ZiShiYing_W(ZiShiYing_N) = mcontrol.Width / ScaleWidth
End If
Next
On Error GoTo 0
'正式

  Open ".\config\BGM.dat" For Input As #1
     Line Input #1, BGMop
  Close #1

  Open ".\config\user.dat" For Input As #2
     Line Input #2, USER
  Close #2
If BGMop = 1 Then
bgm.Show
End If


















End Sub
Private Sub Form_Resize()
Dim mcontrol As Object
On Error Resume Next
ZiShiYing_N = 0
For Each mcontrol In Me.Controls
If Not (TypeOf mcontrol Is Menu) Then
ZiShiYing_N = ZiShiYing_N + 1
mcontrol.Top = ZiShiYing_T(ZiShiYing_N) * ScaleHeight
mcontrol.Left = ZiShiYing_L(ZiShiYing_N) * ScaleWidth
mcontrol.Height = ZiShiYing_H(ZiShiYing_N) * ScaleHeight
mcontrol.Width = ZiShiYing_W(ZiShiYing_N) * ScaleWidth
End If
Next
On Error GoTo 0
End Sub
Private Sub Command1_Click()
End
End Sub

