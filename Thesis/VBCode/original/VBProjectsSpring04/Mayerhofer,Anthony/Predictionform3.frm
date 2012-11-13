VERSION 5.00
Begin VB.Form Predictionform3 
   Caption         =   "Form2"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9885
   LinkTopic       =   "Form2"
   ScaleHeight     =   7215
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton rebounder 
      BackColor       =   &H00C000C0&
      Caption         =   "Display if Player selected is the best rebounder"
      Height          =   1095
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4320
      Width           =   1335
   End
   Begin VB.OptionButton Kevin 
      BackColor       =   &H00FF0000&
      Caption         =   "Kevin"
      Height          =   495
      Left            =   7920
      TabIndex        =   11
      Top             =   2400
      Width           =   1215
   End
   Begin VB.OptionButton Magic 
      BackColor       =   &H0000FFFF&
      Caption         =   "Magic"
      Height          =   495
      Left            =   6120
      TabIndex        =   10
      Top             =   2400
      Width           =   1215
   End
   Begin VB.OptionButton Kareem 
      BackColor       =   &H0000FFFF&
      Caption         =   "Kareem"
      Height          =   495
      Left            =   4320
      TabIndex        =   9
      Top             =   2400
      Width           =   1335
   End
   Begin VB.OptionButton Shaq 
      BackColor       =   &H0000FFFF&
      Caption         =   "Shaquelle"
      Height          =   495
      Left            =   2280
      TabIndex        =   8
      Top             =   2400
      Width           =   1335
   End
   Begin VB.PictureBox picresults6 
      BackColor       =   &H008080FF&
      Height          =   2175
      Left            =   720
      ScaleHeight     =   2115
      ScaleWidth      =   6435
      TabIndex        =   7
      Top             =   4800
      Width           =   6495
   End
   Begin VB.OptionButton Michael 
      BackColor       =   &H000000FF&
      Caption         =   "Michael"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton followingform 
      BackColor       =   &H0080C0FF&
      Caption         =   "Click to Move to Next Form"
      Height          =   975
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5880
      Width           =   1335
   End
   Begin VB.PictureBox picresults1 
      Height          =   1575
      Left            =   240
      Picture         =   "Predictionform3.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox picresults2 
      Height          =   1575
      Left            =   2160
      Picture         =   "Predictionform3.frx":4638
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox picresults3 
      Height          =   1575
      Left            =   4200
      Picture         =   "Predictionform3.frx":9376
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox picresults4 
      Height          =   1575
      Left            =   5880
      Picture         =   "Predictionform3.frx":A99A
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox picresults5 
      BackColor       =   &H00FF0000&
      Height          =   1575
      Left            =   7920
      Picture         =   "Predictionform3.frx":C04B
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label instructions 
      BackColor       =   &H000080FF&
      Caption         =   $"Predictionform3.frx":DA1E
      Height          =   1215
      Left            =   1920
      TabIndex        =   18
      Top             =   3240
      Width           =   4695
   End
   Begin VB.Label player1 
      BackColor       =   &H000000FF&
      Caption         =   "Michael Jordan"
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label player2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Shaquelle O'Neal"
      Height          =   375
      Left            =   2160
      TabIndex        =   16
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label player3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Kareem Abdul Jabbar"
      Height          =   375
      Left            =   4080
      TabIndex        =   15
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label player4 
      BackColor       =   &H0000FFFF&
      Caption         =   "Magic Johnson"
      Height          =   375
      Left            =   6000
      TabIndex        =   14
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label player5 
      BackColor       =   &H00FF0000&
      Caption         =   "Kevin Garnett"
      Height          =   375
      Left            =   7920
      TabIndex        =   13
      Top             =   1920
      Width           =   1455
   End
End
Attribute VB_Name = "Predictionform3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
