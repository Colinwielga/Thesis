VERSION 5.00
Begin VB.Form frmMissionTrips 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Mission Trips (Megan Fitzgerald)"
   ClientHeight    =   6705
   ClientLeft      =   4275
   ClientTop       =   2580
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   9015
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FF8080&
      Caption         =   "Return to Homepage"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5640
      UseMaskColor    =   -1  'True
      Width           =   2175
   End
   Begin VB.CommandButton cmdUserInfo 
      BackColor       =   &H00FF8080&
      Caption         =   "Join Us on a Mission Trip"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      UseMaskColor    =   -1  'True
      Width           =   2775
   End
   Begin VB.CommandButton cmdPictures 
      BackColor       =   &H00FF8080&
      Caption         =   " Pictures from Our Trips"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      UseMaskColor    =   -1  'True
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   120
      Picture         =   "MissionTrips.frx":0000
      ScaleHeight     =   1875
      ScaleWidth      =   3195
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   $"MissionTrips.frx":1BAC2
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   8655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Mission Trips"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   1
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmMissionTrips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjectAmigosForChrist (Megan Fitzgerald's Visual Basic Project.vbp)
'Form Name : frmMissionTrips (MissionTrips.frm)
'Author: Megan Fitzgerald
'Date Written: October 28, 2003
'Purpose of This Form: The purpose of this form is to allow the user to obtain more
                        'information about the mission trips that Amigos for Christ
                        'takes to Nicaragua every year.  The buttons on this form
                        'will allow the user to navigate to other forms to
                        'see pictures from the trips and send contact
                        'infomation to Amigos for Christ.
                        
                    
Option Explicit
'The following buttons will allow the user to navigate between various forms.
'"Hide" means that the form will disappear when the button is clicked
'and "Show means that the form will appear for the user to view.



Private Sub cmdPictures_Click()

frmMissionTrips.Hide
frmPictures.Show

End Sub

Private Sub cmdReturn_Click()

'Take the user back to the Homepage "Amigos for Christ"
frmMissionTrips.Hide
frmHomepage.Show

End Sub

Private Sub cmdUserInfo_Click()

frmMissionTrips.Hide
frmUserInfo.Show

End Sub

