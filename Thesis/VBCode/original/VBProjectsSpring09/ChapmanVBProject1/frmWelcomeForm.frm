VERSION 5.00
Begin VB.Form WelcomeForm 
   Caption         =   "Form1"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   5775
      Left            =   0
      Picture         =   "WelcomeForm.frx":0000
      ScaleHeight     =   5715
      ScaleWidth      =   8235
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin VB.CommandButton cmdConference 
         Caption         =   "Our Conference"
         Height          =   855
         Left            =   6000
         TabIndex        =   3
         Top             =   600
         Width           =   2175
      End
      Begin VB.CommandButton cmdTeam 
         BackColor       =   &H8000000D&
         Caption         =   "Team Info"
         Height          =   855
         Left            =   120
         MaskColor       =   &H00FF0000&
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lblWelcome 
         BackColor       =   &H80000002&
         Caption         =   "Welcome to St. John's Water Polo Club!"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   600
         TabIndex        =   1
         Top             =   120
         Width           =   6015
      End
   End
End
Attribute VB_Name = "WelcomeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdTeam_Click()

frmTeamForm.Show
frmWelcomeForm.Hide

End Sub

Private Sub cmdConference_Click()

frmConferenceForm.Show
frmWelcomeForm.Hide

End Sub
