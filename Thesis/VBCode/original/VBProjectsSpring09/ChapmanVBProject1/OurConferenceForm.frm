VERSION 5.00
Begin VB.Form OurConferenceForm 
   BackColor       =   &H0000FFFF&
   Caption         =   "Form1"
   ClientHeight    =   9525
   ClientLeft      =   1650
   ClientTop       =   1245
   ClientWidth     =   12105
   FillColor       =   &H0000FF00&
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   MousePointer    =   3  'I-Beam
   ScaleHeight     =   9525
   ScaleWidth      =   12105
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to Welcome Page"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7320
      Width           =   2895
   End
   Begin VB.CommandButton cmdTournaments 
      BackColor       =   &H0000FF00&
      Caption         =   "Tournaments"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7800
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      UseMaskColor    =   -1  'True
      Width           =   3855
   End
   Begin VB.CommandButton cmdTeams 
      BackColor       =   &H00FF00FF&
      Caption         =   "Different Teams"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   2520
      Picture         =   "OurConferenceForm.frx":0000
      Top             =   2760
      Width           =   6750
   End
   Begin VB.Label lblHeartland 
      BackColor       =   &H0000FFFF&
      Caption         =   "Heartland Conference"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4320
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
End
Attribute VB_Name = "OurConferenceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WaterPoloProject
'OurConferenceForm
'Bobby Chapman
'Written 3/16/2009
'Objective-provide a page to go to either the DifferentTeamsForm or TournamentsForm
Option Explicit

Private Sub cmdBack_Click()
'goes back to the WelcomeForm
OurConferenceForm.Hide
WelcomeForm.Show

End Sub

Private Sub cmdTeams_Click()
'goes to the DifferentTeamsForm
OurConferenceForm.Hide
DifferentTeamsForm.Show

End Sub

Private Sub cmdTournaments_Click()
'goes to the TournamentForm
OurConferenceForm.Hide
TournamentForm.Show

End Sub
