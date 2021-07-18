VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form logon 
   BorderStyle     =   0  'None
   Caption         =   "你好快啊"
   ClientHeight    =   7215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   7000
      Left            =   1080
      Top             =   2160
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "单击屏幕跳过"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   7215
      Left            =   1080
      TabIndex        =   1
      Top             =   0
      Width           =   10995
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   7680
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10240
      URL             =   ".\videos\startup.mp4"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "none"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   18071
      _cy             =   13547
   End
End
Attribute VB_Name = "logon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Label1_Click()
Unload Me
mainmenu.Show
End Sub

Private Sub Timer1_Timer()

mainmenu.Show
Unload Me
End Sub

Private Sub WindowsMediaPlayer1_Click(ByVal nButton As Integer, ByVal nShiftState As Integer, ByVal fX As Long, ByVal fY As Long)
WindowsMediaPlayer1.Controls.play
End Sub

