VERSION 5.00
Begin VB.Form gmsrt 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7080
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   3840
      Top             =   600
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "老甘正带着56个婊仔赶来..."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   3960
      Width           =   3255
   End
   Begin VB.Label lblWarning 
      Caption         =   "警告：未经允许严禁用于商业用途"
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3045
   End
   Begin VB.Label lblProductName 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "婊仔2019"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   72
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   6300
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "版本"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5160
      TabIndex        =   0
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "gmsrt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public money As Integer
Public X As Integer
Public Y As Integer
Option Explicit
Dim ZiShiYing_T(), ZiShiYing_L(), ZiShiYing_H(), ZiShiYing_W(), ZiShiYing_N%

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
    lblVersion.Caption = "版本 " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
Dim temp As String
  Open ".\game\X.dat" For Input As #3
     Line Input #3, temp
  Close #3
  gmsrt.X = temp '将temp输入到x
    Open ".\game\Y.dat" For Input As #4
     Line Input #4, temp
  Close #4
  gmsrt.Y = temp '将temp输入到y
      Open ".\game\money.dat" For Input As #5
     Line Input #5, temp
  Close #5
  gmsrt.money = temp '将temp输入到money
    
    
    
    
    
    
    
    
  










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

Private Sub Timer1_Timer()
start.Show
Me.Hide
End Sub
