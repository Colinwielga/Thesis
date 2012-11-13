VERSION 5.00
Begin VB.Form TeamForm 
   Caption         =   "Form2"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9195
   LinkTopic       =   "Form2"
   ScaleHeight     =   7215
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCaptains 
      Caption         =   "The Captains"
      Height          =   1455
      Left            =   3120
      TabIndex        =   2
      Top             =   3720
      Width           =   2895
   End
   Begin VB.CommandButton cmdHonors 
      Caption         =   "Honors"
      Height          =   1455
      Left            =   5520
      TabIndex        =   1
      Top             =   1080
      Width           =   3015
   End
   Begin VB.CommandButton cmdMeetTeams 
      Caption         =   "Meet the team"
      Height          =   1455
      Left            =   720
      TabIndex        =   0
      Top             =   1080
      Width           =   2895
   End
End
Attribute VB_Name = "TeamForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdMeetTeams_Click()
frmTeamForm.Hide
frmMeetTheTeamsForm.Show
End Sub

Private Sub cmdHonors_Click()
frmTeamForm.Hide
frmHonorsForm.Show
End Sub

Private Sub cmdCaptains_Click()
frmTeamForm.Hide
frmHistoryForm.Show
End Sub
