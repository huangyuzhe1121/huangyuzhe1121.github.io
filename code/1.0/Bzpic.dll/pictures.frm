VERSION 5.00
Begin VB.Form pictures 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox gym 
      Height          =   375
      Left            =   360
      Picture         =   "pictures.frx":0000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox hyz 
      Height          =   375
      Left            =   0
      Picture         =   "pictures.frx":63342
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "pictures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Image1.Picture = Picture1.Picture
End Sub
