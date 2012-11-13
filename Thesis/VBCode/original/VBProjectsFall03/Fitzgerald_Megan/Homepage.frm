VERSION 5.00
Begin VB.Form frmHomepage 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Homepage (Megan Fitzgerald)"
   ClientHeight    =   7125
   ClientLeft      =   3630
   ClientTop       =   2370
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   9120
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "Donations"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5280
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FF8080&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6480
      Width           =   1455
   End
   Begin VB.PictureBox imgChildInDump 
      Height          =   2175
      Left            =   5640
      Picture         =   "Homepage.frx":0000
      ScaleHeight     =   2115
      ScaleWidth      =   3195
      TabIndex        =   5
      Top             =   2160
      Width           =   3255
   End
   Begin VB.CommandButton cmdOurStory 
      BackColor       =   &H00FF8080&
      Caption         =   "Our Story"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5280
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.CommandButton cmdMissionTrips 
      BackColor       =   &H00FF8080&
      Caption         =   "Mission Trips"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5280
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.CommandButton cmdMissionStatement 
      BackColor       =   &H00FF8080&
      Caption         =   "Mission Statement"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   $"Homepage.frx":1BAC2
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
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   5295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Serving His Poor..."
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Amigos For Christ"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   7935
   End
End
Attribute VB_Name = "frmHomepage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjectAmigosForChrist (Megan Fitzgerald's Visual Basic Project.vbp)
'Form Name : frmHomepage (Homepage.frm)
'Author: Megan Fitzgerald
'Date Written: October 28, 2003
'Purpose of Project: The overall pupose of this project is to inform
                    'the user about the non-profit organization
                    'Amigos for Christ.  This project will serve as
                    'a promotional effort to get people interested
                    'in this organization and it will give them information
                    'about the impoverished people of Nicaragua.
                    'This project will also allow the user to request further
                    'infomation about Amigos for Christ by sending contact
                    'information to a database that can later be accessed by the
                    'creator of this project.
'Purpose of This Form:   The purpose of this form is to act as a "homepage" in which
                        'the user can click on various buttons to obtain infomation
                        'about Amigos for Christ in several specific categories.
                 
'Option Explicit is a command which forces
'the user to explicitly declare all variables
'before they can be used.  It serves as a
'"spell checker" for the form."

Option Explicit
'Clicking on each of the following buttons will take the user to a new form
'in which they can access more specific information about Amigos for Christ.
'"Hide" means that the Homepage will no longer be viewed by user, because
'"Show" enables another form to be accessed.


Private Sub cmdMissionStatement_Click()

frmHomepage.Hide
frmMissionStatement.Show

End Sub

Private Sub cmdMissionTrips_Click()

frmHomepage.Hide
frmMissionTrips.Show
End Sub

Private Sub cmdOurStory_Click()

frmHomepage.Hide
frmOurStory.Show

End Sub

'Allows the user to exit the project.
Private Sub cmdQuit_Click()
End
End Sub

Private Sub Command1_Click()

frmHomepage.Hide
frmDonations.Show
End Sub
