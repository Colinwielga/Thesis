VERSION 5.00
Begin VB.Form ProjectFace1 
   BackColor       =   &H80000001&
   Caption         =   "Minnesota Timberwolves Roster 2003-2004 Season"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      BackColor       =   &H80000001&
      Height          =   1215
      Left            =   360
      Picture         =   "ProjectFace1.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   1035
      TabIndex        =   8
      Top             =   6960
      Width           =   1095
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000001&
      Height          =   1215
      Left            =   9120
      Picture         =   "ProjectFace1.frx":3DE3
      ScaleHeight     =   1155
      ScaleWidth      =   1035
      TabIndex        =   7
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Player Statistics"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   7560
      TabIndex        =   5
      Top             =   5880
      Width           =   2175
   End
   Begin VB.PictureBox Picture2 
      Height          =   2175
      Left            =   8280
      Picture         =   "ProjectFace1.frx":7BC6
      ScaleHeight     =   2115
      ScaleWidth      =   2115
      TabIndex        =   4
      Top             =   720
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   120
      Picture         =   "ProjectFace1.frx":102A1
      ScaleHeight     =   2115
      ScaleWidth      =   2115
      TabIndex        =   3
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton cmdRoster 
      Caption         =   "Load 2003-2004 Team Roster"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   5880
      Width           =   2175
   End
   Begin VB.PictureBox picresults 
      Height          =   3735
      Left            =   2400
      ScaleHeight     =   3675
      ScaleWidth      =   5715
      TabIndex        =   1
      Top             =   960
      Width           =   5775
   End
   Begin VB.Label Label2 
      Caption         =   "Designer: Ben Lewis"
      Height          =   255
      Left            =   8760
      TabIndex        =   9
      Top             =   8280
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "             Minnesota Timberwolves 2003 - 2004                    MINNESOTA TIMBERWOLVES 2003-2004"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "ProjectFace1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Minnesota Timberwolves Basketball Season 2003-2004
'Project 1 (Project1.vbp)
'ProjectFace1 (ProjectFace1.Frm)
'Ben Lewis
'March 13, 2004
'The Purpose of this project is to gain knowledge of the statistics from various players of the Minnesota Basketball team for 2003-2004 season
'The Purpose of this form is to visually show the user what the program executes
'Also, this form is a guide for debugging problems with variables, and making sure all variables used in computations are declared
Option Explicit




Private Sub cmdLoad_Click()
'This command will load all of the individual statistics
'This command will load Statistics page that arranges statistics in various ways (i.e. players, numbers, points, rebounds, etc.)
ProjectFace1.Hide
ProjectFace2.Show
cmdRoster.Enabled = True
cmdLoad.Enabled = True
End Sub




Private Sub cmdQuit_Click()
    'Ends the program if the user wishes to Stop
    End
End Sub





Private Sub cmdRoster_Click()
'This command will display player names and numbers that are on the current roster
'This command will be helpful for user to see players and the spelling of their names, which they will have to use in inputbox of favorite player
'This clears the picturebox of any input (if any)
picresults.Cls
picresults.Print "Minnesota Timberwolves 2003-2004 Roster"
picresults.Print
picresults.Print "           Player                                                                       #"
picresults.Print "*************************************************************************************************"
'Open "N:\CS130\handin\Lewis, Ben\roster.txt" For Input As #1
Open PATH & "roster.txt" For Input As #1
    For I = 1 To 13
        Input #1, Names(I), Numbers(I)
        picresults.Print Tab(5); Names(I); Tab(55); Numbers(I)
    Next I
Close
cmdLoad.Enabled = True
cmdRoster.Enabled = False
End Sub

