VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11160
   LinkTopic       =   "Form2"
   ScaleHeight     =   9195
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   6
      Top             =   8280
      Width           =   1815
   End
   Begin VB.CommandButton cmdPunchBunch 
      Caption         =   "Punch a Bunch"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1320
      TabIndex        =   3
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CommandButton cmd5Chances 
      Caption         =   "Five Chances"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1320
      TabIndex        =   2
      Top             =   6360
      Width           =   2415
   End
   Begin VB.CommandButton cmdEasy123 
      Caption         =   "Eazy as 1-2-3"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1320
      TabIndex        =   1
      Top             =   2400
      Width           =   2415
   End
   Begin VB.PictureBox picBob 
      BorderStyle     =   0  'None
      Height          =   7455
      Left            =   4920
      Picture         =   "Project2.frx":0000
      ScaleHeight     =   7455
      ScaleWidth      =   6015
      TabIndex        =   0
      Top             =   480
      Width           =   6015
   End
   Begin VB.Label lblChoose 
      BackColor       =   &H80000008&
      Caption         =   "Please choose a pricing game:"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   4575
   End
   Begin VB.Label lblCongrats 
      BackColor       =   &H80000008&
      Caption         =   "Contratulations!"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   840
      TabIndex        =   4
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd5Chances_Click()  'This screen is simply used for the player to select which of the pricing games they would like to play.
Form5.Show  'This closes Form2 and opens Form5 which is the Five Chances game.
Form2.Hide
End Sub

Private Sub cmdEasy123_Click()
Form3.Show  'This closes Form2 and opens Form3 which is the Easy as 1-2-3 game.
Form2.Hide
End Sub

Private Sub cmdPunchBunch_Click()
Form4.Show  'This closes Form2 and opens Form4 which is the Punch a Bunch game.
Form2.Hide
End Sub

Private Sub Command1_Click()    'This button is used if the player decides to quit the game at this point.
End
End Sub
