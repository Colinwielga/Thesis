VERSION 5.00
Begin VB.Form TeamForm 
   BackColor       =   &H8000000D&
   Caption         =   "Form2"
   ClientHeight    =   9675
   ClientLeft      =   1770
   ClientTop       =   1125
   ClientWidth     =   11970
   LinkTopic       =   "Form2"
   MousePointer    =   7  'Size N S
   ScaleHeight     =   9675
   ScaleWidth      =   11970
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF0000&
      Caption         =   "Back To Main"
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7800
      Width           =   3015
   End
   Begin VB.CommandButton cmdCaptains 
      BackColor       =   &H00808000&
      Caption         =   "The Captains"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7920
      Width           =   3015
   End
   Begin VB.CommandButton cmdHonors 
      BackColor       =   &H0000FFFF&
      Caption         =   "Honors"
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton cmdMeetTeams 
      BackColor       =   &H00004080&
      Caption         =   "Meet the team"
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
      Left            =   240
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   6915
      Left            =   2040
      Picture         =   "TeamForm.frx":0000
      Top             =   1320
      Width           =   7500
   End
End
Attribute VB_Name = "TeamForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WaterPoloProject
'TeamForm
'Bobby Chapman
'Written 3/14/2009
'Objective-provide a page to go to either the MeetTheTeamForm, TheCaptainsForm,
'the HonorsForm, or to the WelcomeForm
Option Explicit

Private Sub cmdMeetTeams_Click()
'goes the the MeetTheTeamForm
TeamForm.Hide
MeetTheTeamForm.Show
End Sub

Private Sub cmdHonors_Click()
'goes to the HonorsForm
TeamForm.Hide
HonorsForm.Show
End Sub

Private Sub cmdCaptains_Click()
'goes to TheCaptainsForm
TeamForm.Hide
TheCaptainsForm.Show
End Sub

Private Sub cmdBack_Click()
'goes back to the WelcomeForm
TeamForm.Hide
WelcomeForm.Show
End Sub

