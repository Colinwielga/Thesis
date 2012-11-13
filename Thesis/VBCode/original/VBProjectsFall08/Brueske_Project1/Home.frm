VERSION 5.00
Begin VB.Form Home 
   BackColor       =   &H8000000D&
   Caption         =   "Home"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   3000
      TabIndex        =   11
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H0080FFFF&
      Caption         =   "Other"
      Height          =   2055
      Left            =   2760
      TabIndex        =   9
      Top             =   3360
      Width           =   1935
      Begin VB.CommandButton cmdbg 
         Caption         =   "Info"
         Height          =   615
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Results"
      Height          =   2055
      Left            =   5400
      TabIndex        =   7
      Top             =   1080
      Width           =   1935
      Begin VB.CommandButton cmdresults 
         Caption         =   "User Results"
         Enabled         =   0   'False
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame frame2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Tests"
      Height          =   2055
      Left            =   2760
      TabIndex        =   4
      Top             =   1080
      Width           =   1935
      Begin VB.CommandButton cmdalienation 
         Caption         =   "Alienation Test"
         Enabled         =   0   'False
         Height          =   615
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmddistance 
         Caption         =   "Social Distance Test"
         Enabled         =   0   'False
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Data"
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
      Begin VB.CommandButton cmduser 
         Caption         =   "Create Profile"
         Height          =   615
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdLoad 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Load Profiles"
         Height          =   615
         Left            =   240
         MaskColor       =   &H00C0C0FF&
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Alienation and Social Distance Project"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "Home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'  Alienation and Social Distance Project
'Home Form
'Kevin Brueske
'Created Oct 26, 2008
'Objective
    'This form will act as a gateway to the other forms. It will also allow the user to load data and create a user profile
    'which is necessary for using the other parts of the program.
'Projective Objective
    'In all,the objective of this program is to determine and compare alienation and social distance scores
    'of users. This program will accomplish this by loading data (for comparison),
    'allowing a user to create a user profile, taking the alienation and social distance test, and then comparing the
    'results of the tests with other preloaded or generated users.




   

Private Sub cmdalienation_Click()
     'User profile mechanism
     'If an alienation score exists for a certain user, then that user has to create
     'a new user profile.
     'Otherwise, the user is allowed to take the test
     If alienationScore(usrnum) > 0 Then
        MsgBox "Create New User"
        Home.Show
    Else:
    Home.Hide
    Alienation.Show
    End If
End Sub

Private Sub cmdbg_Click()
     Information.Show
    Home.Hide
End Sub

Private Sub cmddistance_Click()
    'User profile mechanism
     'If a social distance score exists for a certain user, then that user has to create
     'a new user profile.
     'Otherwise, the user is allowed to take the test
    If distanceScore(usrnum) > 0 Then
    MsgBox "Create New User"
    Home.Show
    Else:
    Home.Hide
    SocialDistance.Show
    End If
End Sub

Private Sub cmdLoad_Click()
    'File Input, load data into project wide data
    Open App.Path & "\people.txt" For Input As #1
        Do Until EOF(1)
            usrnum = usrnum + 1
            Input #1, fname(usrnum), lname(usrnum), age(usrnum), major(usrnum), socialclass(usrnum), religion(usrnum), alienationScore(usrnum), distanceScore(usrnum)
        Loop
        Close #1
        
    'After completing the load, the load button will be disabled and the data button will be enabled
    cmdresults.Enabled = True
    cmdLoad.Enabled = False
End Sub

Private Sub cmdQuit_Click()
    'end
    End
End Sub

Private Sub cmdresults_Click()
    'Switching from the home form to the results form
    Home.Hide
    Results.Show
End Sub

Private Sub cmduser_Click()
    'keep track of the total number of users
    usrnum = usrnum + 1
    'Input user profile data
    fname(usrnum) = InputBox("Enter your first name.", "Name")
    lname(usrnum) = InputBox("Enter your last name.", "Name")
    age(usrnum) = InputBox("Enter your age.", "Age")
    major(usrnum) = InputBox("Enter your major.", "Major")
    socialclass(usrnum) = InputBox("Enter your social class (Lower, Middle, Upper).", "Social Class")
    religion(usrnum) = InputBox("Enter your religion.", "Religion")
    
    'Enable the other buttons now that a user profile has been created
    cmdresults.Enabled = True
    cmddistance.Enabled = True
    cmdAlienation.Enabled = True
    
    
End Sub




