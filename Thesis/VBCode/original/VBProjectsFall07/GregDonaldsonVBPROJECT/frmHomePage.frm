VERSION 5.00
Begin VB.Form frmHomePage 
   BackColor       =   &H00FF0000&
   Caption         =   "Home Page"
   ClientHeight    =   5130
   ClientLeft      =   1770
   ClientTop       =   1365
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   7980
   Begin VB.CommandButton cmdNewTime 
      Caption         =   "Compare New Time"
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdMeetTeam 
      Caption         =   "Meet The Team"
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   2760
      Width           =   975
   End
   Begin VB.PictureBox picsju 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FF0000&
      Height          =   975
      Left            =   3240
      Picture         =   "frmHomePage.frx":0000
      ScaleHeight     =   975
      ScaleWidth      =   1575
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdSearchSwimmer 
      BackColor       =   &H000000FF&
      Caption         =   "Search Swimmers"
      Height          =   495
      Left            =   4920
      MaskColor       =   &H000000FF&
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdBestTimes 
      Caption         =   "Best Times"
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblsju 
      BackColor       =   &H00FF0000&
      Caption         =   "Saint John's University Swimming"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   3720
      Width           =   2415
   End
End
Attribute VB_Name = "frmHomePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'all of the buttons on the home page lead to another form associated with each button.
'the home page acts as a menu

Private Sub cmdBestTimes_Click()
frmHomePage.Hide
frmBestTimes.Show

End Sub

Private Sub cmdMeetTeam_Click()

frmHomePage.Hide
frmRoster.Show

End Sub

Private Sub cmdNewTime_Click()
frmHomePage.Hide
frmNewTime.Show

End Sub

Private Sub cmdSearchSwimmer_Click()

frmHomePage.Hide
frmIndividualSearch.Show

End Sub
