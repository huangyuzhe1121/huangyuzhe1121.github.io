VERSION 5.00
Begin VB.Form setting 
   BorderStyle     =   0  'None
   Caption         =   "����"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   120
      TabIndex        =   4
      Text            =   "��֪�����"
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���沢�˳�"
      Default         =   -1  'True
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�ر�"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Line Line3 
      X1              =   3360
      X2              =   3360
      Y1              =   0
      Y2              =   5880
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   3360
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3360
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label1 
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   27.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "�û���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   27.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   2295
   End
End
Attribute VB_Name = "setting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Sub Command1_Click()
mainmenu.BGMop = 1
bgm.Hide

End Sub

Private Sub Command2_Click()
mainmenu.BGMop = 0
Unload bgm

End Sub

Private Sub Command3_Click()
Dim temp As String

temp = Text1.Text

 Open ".\config\BGM.dat" For Output As #1
 Print #1, mainmenu.BGMop
 Close #1
 
 Open ".\config\USER.dat" For Output As #2
 Print #2, temp
 Close #2

mainmenu.USER = temp

Unload Me
End Sub

Private Sub Command4_Click()
Text1.Text = ""
End Sub

Private Sub Form_Load()
Text1.Text = mainmenu.USER
  SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

