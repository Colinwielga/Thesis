VERSION 5.00
Begin VB.Form Player 
   BackColor       =   &H000080FF&
   Caption         =   "Choose a player!"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4800
      TabIndex        =   5
      Text            =   "Tanya"
      Top             =   6360
      Width           =   855
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3000
      TabIndex        =   4
      Text            =   "Jason"
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7200
      TabIndex        =   3
      Text            =   "Charlie"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      TabIndex        =   2
      Text            =   "Sheryl"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Text            =   "Ned"
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox Txtplayer 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   0
      Text            =   "Pick an owner..."
      Top             =   360
      Width           =   3255
   End
   Begin VB.Image charlie1 
      Height          =   2220
      Left            =   7080
      Picture         =   "Player.frx":0000
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1530
   End
   Begin VB.Image jason1 
      Height          =   2775
      Left            =   720
      Picture         =   "Player.frx":4A07
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Image sheryl1 
      Height          =   2175
      Left            =   3480
      Picture         =   "Player.frx":7A80
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Image tanya1 
      Height          =   2775
      Left            =   5880
      Picture         =   "Player.frx":F49E
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   2655
   End
   Begin VB.Image ned1 
      Height          =   2220
      Left            =   720
      Picture         =   "Player.frx":26F56
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1635
   End
End
Attribute VB_Name = "Player"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub charlie1_Click()
Player.Hide
charlie.Show 'sends a player to the approriet form/profile and set a unique number to a counter so you always know which person it is
picpick = 3
End Sub

Private Sub jason1_Click()
Player.Hide
Jason.Show 'sends a player to the approriet form/profile and set a unique number to a counter so you always know which person it is
picpick = 4
End Sub

Private Sub Ned1_Click()
Player.Hide 'sends a player to the approriet form/profile and set a unique number to a counter so you always know which person it is
Ned.Show
picpick = 1
End Sub

Private Sub sheryl1_Click()
Player.Hide 'sends a player to the approriet form/profile and set a unique number to a counter so you always know which person it is
sheryl.Show
picpick = 2
End Sub

Private Sub tanya1_Click()
Player.Hide
tanya.Show 'sends a player to the approriet form/profile and set a unique number to a counter so you always know which person it is
picpick = 5
End Sub

