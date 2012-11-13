VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H00FF0000&
   Caption         =   "Menu"
   ClientHeight    =   6750
   ClientLeft      =   3240
   ClientTop       =   2655
   ClientWidth     =   11070
   LinkTopic       =   "Form2"
   ScaleHeight     =   6750
   ScaleWidth      =   11070
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Quit"
      Height          =   735
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton cmdHistory 
      BackColor       =   &H000000FF&
      Caption         =   "History"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton cmdWhatsRugby 
      BackColor       =   &H8000000A&
      Caption         =   "Click here for more information about rugby!"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Width           =   4575
   End
   Begin VB.CommandButton cmdSchedule 
      BackColor       =   &H000000FF&
      Caption         =   "Schedules and Recent Scores"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdStats 
      BackColor       =   &H000000FF&
      Caption         =   "Stats"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdPlayers 
      BackColor       =   &H000000FF&
      Caption         =   "Player Profiles"
      BeginProperty Font 
         Name            =   "Myriad Web Pro Condensed"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.PictureBox picSlideShow 
      Height          =   5775
      Left            =   120
      Picture         =   "frmMenu.frx":0000
      ScaleHeight     =   5715
      ScaleWidth      =   8955
      TabIndex        =   0
      Top             =   120
      Width           =   9015
   End
   Begin VB.Label lblText 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "                                     Click picture for slideshow"
      Height          =   555
      Left            =   960
      TabIndex        =   7
      Top             =   6000
      Width           =   1950
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'St. John's Rugby
'Sam Herrmann
'March 2009
'This form acts as a navigation menu for all the other forms in the project

'It uses command buttons that bring up other forms while hiding the menu
Option Explicit
Private Sub cmdHistory_Click()

frmMenu.Hide
frmHistory.Show

End Sub

Private Sub cmdPlayers_Click()

frmMenu.Hide
frmProfiles.Show

End Sub

Private Sub cmdSchedule_Click()

frmMenu.Hide
frmSchedulesScores.Show

End Sub

Private Sub cmdStats_Click()

frmMenu.Hide
frmStats.Show

End Sub

Private Sub cmdWhatsRugby_Click()

frmMenu.Hide
frmWhatsRugby.Show

End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub picSlideShow_Click()

Dim names(1 To 5) As String
Dim choosePic As Integer, stopper As Integer
Dim t As Double, lastPic As Integer, CTR As Double

names(1) = "teamIA.jpg"
names(2) = "jackethug.jpg"
names(3) = "lineoutUofM.jpg"
names(4) = "docRun2.jpg"
names(5) = "teamAllmn.jpg"

choosePic = 1
stopper = 0

Do While (stopper < 5)
    picSlideShow.Picture = LoadPicture(App.Path & "\Pictures\SlideShow\" & names(choosePic))

t = Timer
    Do While (Timer - t) < 1
        CTR = CTR + 1
        If CTR = 1000000 Then
            CTR = 0
        End If
    Loop
        stopper = stopper + 1
        lastPic = choosePic
        choosePic = (stopper Mod CTR) + 1
Loop

End Sub

