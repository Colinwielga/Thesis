VERSION 5.00
Begin VB.Form frmMainPage 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Minnesota Timberwolves"
   ClientHeight    =   6735
   ClientLeft      =   3165
   ClientTop       =   2280
   ClientWidth     =   10080
   FillColor       =   &H00800080&
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   10080
   Begin VB.CommandButton cmdQuit 
      Caption         =   "      Leave        (But I wouldn't recomend it)"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8040
      TabIndex        =   10
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton cmdAcknowledgements 
      Caption         =   "Acknowledge-ments"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuiz 
      Caption         =   "Test Your Tmiberwolves Knowledge"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8040
      TabIndex        =   8
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton cmdDunkContes 
      Caption         =   "Cast Your Vote For The Dunk Contest"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8040
      TabIndex        =   7
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdKG 
      Caption         =   "The Greatest Timberwolf EVER"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8040
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdCompare 
      Caption         =   "Compare the T-Wolves to Other NBA Teams"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5160
      TabIndex        =   5
      Top             =   5520
      Width           =   2655
   End
   Begin VB.CommandButton cmdDancer 
      Caption         =   "Win a Date With a Timberwolves Dancer"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2280
      TabIndex        =   4
      Top             =   5520
      Width           =   2415
   End
   Begin VB.CommandButton cmdMakeaTeam 
      Caption         =   "Make Your Own Team"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   3
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdCoachingStaff 
      Caption         =   "Meet the Coaches"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdMeetPlayers 
      BackColor       =   &H8000000D&
      Caption         =   "Meet The Players"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      MaskColor       =   &H0000C000&
      TabIndex        =   0
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblMainPage 
      BackColor       =   &H8000000D&
      Caption         =   "Everything You Ever Wanted To Know About The Minnesota                                              Timberwolves"
      BeginProperty Font 
         Name            =   "Kozuka Gothic Pro B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   9015
   End
   Begin VB.Image Image1 
      Height          =   4215
      Left            =   2280
      Picture         =   "frmMainPage.frx":0000
      Top             =   1080
      Width           =   5520
   End
End
Attribute VB_Name = "frmMainPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This is the main page of my project.  It serves as link to all of the possible options of the program
'There are a total of 10 command buttons on this page.
'Each button will lead the user to a different portion of my program



Private Sub cmdAcknowledgements_Click()
'This button displays a message box which offers acknowledgment for my project
MsgBox "There are a few people I need to thank for the completion of this project.  First of all, the majority of credit goes to myself.  Thanks to my creative genius this project exists.  Second of all, thanks is owed to God who makes everything possible.  And lastly, I'm endowed to the brilliant Imad Rahal who shared with me his extensive VB knowledge"

End Sub

Private Sub cmdCoachingStaff_Click()
'This button leads to a form which gives info about the coaching staff
frmMainPage.Visible = False
frmCoachingStaff.Visible = True

End Sub

Private Sub cmdCompare_Click()
'This button leads to a form which gives info about the timberwolves ranking and the rankings of all NBA teams
frmMainPage.Visible = False
frmCompare.Visible = True

End Sub

Private Sub cmdDancer_Click()
'this button leads to a form which is a "dating" game that matches the user up with a timberwolves dancer or Crunch the mascot
frmMainPage.Visible = False
frmDancer.Visible = True

End Sub

Private Sub cmdDunkContes_Click()
'this button leads the user to a form which is a dunk contest.  The user can vote for their favorite dunk and then display results
frmMainPage.Visible = False
frmDunkContest.Visible = True

End Sub

Private Sub cmdKG_Click()
'this button leads the user to a form which is devoted to Kevin Garnett.  The form gives various info about KG
frmMainPage.Visible = False
frmKG.Visible = True

End Sub

Private Sub cmdMakeaTeam_Click()
'This button leads to a form which allows the user to choose their own timberwolves starting lineup.  I then rank their lineup.
frmMainPage.Visible = False
frmMakeaTeam.Visible = True

End Sub

Private Sub cmdMeetPlayers_Click()

'This button leads to a form which allows the user to see the timberwolves roster and learn information about specific timberwolves players
frmMainPage.Visible = False
frmMeetPlayers.Visible = True

End Sub

Private Sub cmdQuit_Click()
'This button ends my program
End

End Sub

Private Sub cmdQuiz_Click()
'This button leads to a form which gives the user a timberwolves quiz
frmMainPage.Visible = False
frmQuiz.Visible = True

End Sub
