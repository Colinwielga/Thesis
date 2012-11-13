VERSION 5.00
Begin VB.Form bbbb
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   10515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   10515
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1
      Height          =   5775
      Left            =   360
      Picture         =   "bbbb.frx":0000
      ScaleHeight     =   5715
      ScaleWidth      =   7635
      TabIndex        =   7
      Top             =   4560
      Width           =   7695
   End
   Begin VB.CommandButton cmdLast
      Caption         =   "Go to our thank you page"
      BeginProperty Font
         Name            =   "Maiandra GD"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5640
      TabIndex        =   6
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton cmdMusic
      Caption         =   "Find out what music they listen to"
      BeginProperty Font
         Name            =   "Maiandra GD"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3240
      TabIndex        =   5
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton cmdPoll
      Caption         =   "Find Out How Awesome You Are!"
      BeginProperty Font
         Name            =   "Maiandra GD"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   720
      TabIndex        =   4
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton cmdpics
      Caption         =   "Get to know the Packer's"
      BeginProperty Font
         Name            =   "Maiandra GD"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   720
      TabIndex        =   3
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton cmdSwap
      Caption         =   "Get the Player's Stats from the 2009 season."
      BeginProperty Font
         Name            =   "Maiandra GD"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5640
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton cmdData
      Caption         =   "Get the Player's Personal Info"
      BeginProperty Font
         Name            =   "Maiandra GD"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3120
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "What would you like to do?"
      BeginProperty Font
         Name            =   "Vrinda"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
   Begin VB.Shape Shape1
      FillColor       =   &H0000FFFF&
      FillStyle       =   3  'Vertical Line
      Height          =   10575
      Left            =   0
      Top             =   -120
      Width           =   10695
   End
End
Attribute VB_Name = "bbbb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Get to know the Packers' Receivers
'bbbb
'Sam Pederson
'2/18/10
'This form is the menu form that will take you to where you want to go
Private Sub cmdData_Click() 'this button takes you to cccc
    aaaa.Hide
    bbbb.Hide
    cccc.Hide
    dddd.Show
    eeee.Hide
    ffff.Hide
    gggg.Hide
    hhhh.Hide
End Sub

Private Sub cmdLast_Click() 'this button takes you to hhhh
    aaaa.Hide
    bbbb.Hide
    cccc.Hide
    dddd.Hide
    eeee.Hide
    ffff.Hide
    gggg.Hide
    hhhh.Show
End Sub

Private Sub cmdMusic_Click() 'this button takes you to gggg
    aaaa.Hide
    bbbb.Hide
    cccc.Hide
    dddd.Hide
    eeee.Hide
    ffff.Hide
    gggg.Show
    hhhh.Hide
End Sub

Private Sub cmdpics_Click() 'this button takes you to ffff
    aaaa.Hide
    bbbb.Hide
    cccc.Hide
    dddd.Hide
    eeee.Hide
    ffff.Show
    gggg.Hide
    hhhh.Hide
End Sub

Private Sub cmdPoll_Click() 'this button takes you to cccc
    aaaa.Hide
    bbbb.Hide
    cccc.Show
    dddd.Hide
    eeee.Hide
    ffff.Hide
    gggg.Hide
    hhhh.Hide
End Sub

Private Sub cmdSwap_Click() 'this button takes you to eeee
    aaaa.Hide
    bbbb.Hide
    cccc.Hide
    dddd.Hide
    eeee.Show
    ffff.Hide
    gggg.Hide
    hhhh.Hide
End Sub
