VERSION 5.00
Begin VB.Form Swing 
   BackColor       =   &H000080FF&
   Caption         =   "Form1"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9180
   FillColor       =   &H000080FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Go Back to Main Page"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7080
      TabIndex        =   3
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7080
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   4335
      Left            =   2160
      Picture         =   "Swing.frx":0000
      ScaleHeight     =   4275
      ScaleWidth      =   4275
      TabIndex        =   0
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "You Lose!"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2760
      TabIndex        =   2
      Top             =   4920
      Width           =   3375
   End
End
Attribute VB_Name = "Swing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'http://www.vbforums.com/showthread.php?t=350962
'
Option Explicit
'http://support.microsoft.com/kb/231298
'http://www.thescripts.com/forum/thread14392.html
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Command1_Click()
Dim CTR As Integer
CTR = 5
Do While CTR > 0
    Picture1.Picture = LoadPicture(App.Path & "\midleft.bmp")
    Sleep (200)
    Picture1.Picture = LoadPicture(App.Path & "\left.bmp")
    Sleep (200)
    Picture1.Picture = LoadPicture(App.Path & "\midleft.bmp")
    Sleep (200)
    Picture1.Picture = LoadPicture(App.Path & "\mid.bmp")
    Sleep (200)
    Picture1.Picture = LoadPicture(App.Path & "\midright.bmp")
    Sleep (200)
    Picture1.Picture = LoadPicture(App.Path & "\right.bmp")
    Sleep (200)
    Picture1.Picture = LoadPicture(App.Path & "\midright.bmp")
    Sleep (200)
    Picture1.Picture = LoadPicture(App.Path & "\mid.bmp")
    Sleep (200)
CTR = CTR - 1
Loop
End Sub

Private Sub Command2_Click()
Swing.Hide
frmMainpage.Show
End Sub
