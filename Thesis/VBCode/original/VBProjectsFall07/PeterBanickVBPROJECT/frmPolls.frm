VERSION 5.00
Begin VB.Form frmPolls 
   Caption         =   "What the Consensus?"
   ClientHeight    =   10545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15210
   LinkTopic       =   "Form1"
   ScaleHeight     =   10545
   ScaleWidth      =   15210
   Begin VB.CommandButton cmdPublicOpinion 
      Enabled         =   0   'False
      Height          =   1935
      Left            =   360
      Picture         =   "frmPolls.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6720
      Width           =   2895
   End
   Begin VB.CommandButton cmdYourOpinion 
      Height          =   1935
      Left            =   360
      Picture         =   "frmPolls.frx":A8E1
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   2895
   End
   Begin VB.CommandButton cmdReturnMenu 
      BackColor       =   &H000000FF&
      DisabledPicture =   "frmPolls.frx":1583D
      Height          =   1215
      Left            =   360
      Picture         =   "frmPolls.frx":1D619
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9120
      Width           =   1935
   End
   Begin VB.Image picRoseSlide 
      Height          =   12090
      Left            =   0
      Picture         =   "frmPolls.frx":24F5E
      Top             =   0
      Width           =   15300
   End
End
Attribute VB_Name = "frmPolls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPublicOpinion_Click()
    'shows form PublicOpinion to see what other sports fans have to say about this topic
    frmPublicOpinion.Show
End Sub

Private Sub cmdReturnMenu_Click()
    'returns user to previous screen (Polls) for further use
    cmdYourOpinion.Enabled = True
    frmBio.Hide
    frmMenuPage.Show
End Sub

Private Sub cmdYourOpinion_Click()
    'shows form YourOpinion for user to take a quick survey to see what he/she thinks about the issue
    frmYourOpinion.Show
    cmdPublicOpinion.Enabled = True
    cmdYourOpinion.Enabled = False
End Sub
