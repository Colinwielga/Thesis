VERSION 5.00
Begin VB.Form WelcomeForm 
   Caption         =   "Form1"
   ClientHeight    =   9450
   ClientLeft      =   1545
   ClientTop       =   1125
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   MousePointer    =   14  'Arrow and Question
   ScaleHeight     =   9450
   ScaleWidth      =   11985
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000009&
      Height          =   9375
      Left            =   0
      Picture         =   "WelcomeForm.frx":0000
      ScaleHeight     =   9315
      ScaleWidth      =   11835
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      Begin VB.CommandButton cmdTeam 
         BackColor       =   &H8000000D&
         Caption         =   "Team Info"
         BeginProperty Font 
            Name            =   "Ravie"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   2475
      End
      Begin VB.CommandButton cmdConference 
         BackColor       =   &H0080C0FF&
         Caption         =   "Our Conference"
         BeginProperty Font 
            Name            =   "Ravie"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   2475
      End
      Begin VB.CommandButton cmdQuit 
         BackColor       =   &H000000C0&
         Caption         =   "Quit"
         BeginProperty Font 
            Name            =   "Ravie"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   7320
         Width           =   2475
      End
      Begin VB.Label lblWelcome 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "Welcome to St. John's Water Polo Club!"
         BeginProperty Font 
            Name            =   "Snap ITC"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1095
         Left            =   2880
         TabIndex        =   4
         Top             =   0
         Width           =   5895
      End
   End
End
Attribute VB_Name = "WelcomeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WaterPoloProject
'WelcomeForm
'Bobby Chapman
'Written 3/14/2009
'Objective- This program prints the members of the SJU Water Polo Team,
'let's the user search for someone on the team, arrange them in alphabetical order,
'shows honors different people have received, shows who the captains are,
'prints the tournaments we had last season and in chronological order, and
'prints the different teams in our Conference

Option Explicit

Private Sub cmdTeam_Click()
'goes to the TeamForm
TeamForm.Show
WelcomeForm.Hide

End Sub

Private Sub cmdConference_Click()
'goes to the OurConferenceForm
OurConferenceForm.Show
WelcomeForm.Hide

End Sub

Private Sub cmdQuit_Click()
'quits the program
End
End Sub
