VERSION 5.00
Begin VB.Form frmCoachingStaff 
   Caption         =   "Coaching Staff"
   ClientHeight    =   8445
   ClientLeft      =   2340
   ClientTop       =   1275
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   10590
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Page"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      TabIndex        =   11
      Top             =   7440
      Width           =   3375
   End
   Begin VB.PictureBox picCoach 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   5400
      ScaleHeight     =   4395
      ScaleWidth      =   4275
      TabIndex        =   10
      Top             =   2760
      Width           =   4335
   End
   Begin VB.CommandButton cmdCoach 
      BackColor       =   &H8000000D&
      Caption         =   "Learn About a Specific Coach"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5640
      MaskColor       =   &H0080FFFF&
      TabIndex        =   9
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Image Image2 
      Height          =   3480
      Left            =   960
      Picture         =   "frmCoachingStaff.frx":0000
      Top             =   5160
      Width           =   3750
   End
   Begin VB.Label lblFarnam 
      BackColor       =   &H80000009&
      Caption         =   "Athletic Trainer - Greg Farnam"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   8
      Top             =   4800
      Width           =   4575
   End
   Begin VB.Label lblVitel 
      BackColor       =   &H80000009&
      Caption         =   "Strength and Conditioning - Dave Vitel"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   7
      Top             =   4440
      Width           =   4815
   End
   Begin VB.Label lblHaskins 
      BackColor       =   &H80000009&
      Caption         =   "Advance Scout - Brent Haskins"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   4080
      Width           =   3735
   End
   Begin VB.Label lblPickney 
      BackColor       =   &H80000009&
      Caption         =   "Ed Pickney"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label lblBickerstaff 
      BackColor       =   &H80000009&
      Caption         =   "John Blair-Bickerstaff"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Label lblSichting 
      BackColor       =   &H80000009&
      Caption         =   "Jerry Sichting"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label lblOcipka 
      BackColor       =   &H80000009&
      Caption         =   "Assistant Coaches - Bob Ocipka"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   2640
      Width           =   3735
   End
   Begin VB.Label lblWittman 
      BackColor       =   &H80000009&
      Caption         =   "Head Coach - Randy Wittman"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Label lblCoaches 
      BackColor       =   &H8000000E&
      Caption         =   "The Coaches"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   15780
      Left            =   -240
      Picture         =   "frmCoachingStaff.frx":3874
      Top             =   -1200
      Width           =   11160
   End
End
Attribute VB_Name = "frmCoachingStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub cmdCoach_Click()
'This button provides information about all of the coaches on the Timberwolves coaching staff
'An input box will ask the user to enter the coach who they want to learn about.
'If Coach matches one of the possible coaches names then info will be shown about that coach
Dim Coach As String
Coach = InputBox("Which Coach Do You Want to Learn About?", "Coaching Staff")
'I will use a select case statement because i'm dealing with more than 2 options
Select Case Coach
    Case Is = "Randy Wittman"
        'I must clear the picture box so that the user can enter multiple names
        'This gives info about Wittman
        picCoach.Cls
        picCoach.Print "Randy Wittman is the Head Coach"
        picCoach.Print "of the Minnesota Timberwolves."
        picCoach.Print "Wittman attended the University"
        picCoach.Print "of Indiana where he won the NCAA "
        picCoach.Print "Championship in 1981. Randy"
        picCoach.Print "was a first round draft choice"
        picCoach.Print "and played in the NBA for "
        picCoach.Print "9 years."
    Case Is = "Bob Ocipka"
        picCoach.Cls
        'This gives info about Ocipka
        picCoach.Print "Bob Ocipka is an assistant coach"
        picCoach.Print "for the Minnesota Timberwoves."
        picCoach.Print "Bob attended Quincy College and"
        picCoach.Print "played for the Pistons in the 80's."
        picCoach.Print "Bob first became well-known for"
        picCoach.Print "coaching prep-school."
    Case Is = "Jerry Sichting"
        picCoach.Cls
        'This gives info about Sichting
        picCoach.Print "Jerry Sichting is an assistant coach"
        picCoach.Print "for the Minnesota Timberwolves."
        picCoach.Print "Sichting attended Purdue. Sichting"
        picCoach.Print "played 10 seasons in the NBA and"
        picCoach.Print "was known for his hard-nosed "
        picCoach.Print "tenacious play."
    Case Is = "John Blair-Bickerstaff"
        picCoach.Cls
        'This gives info about Bickerstaff
        picCoach.Print "Bickerstaff is an assistant coach"
        picCoach.Print "for the Minnesota Timberwolves."
        picCoach.Print "Bickerstaff played for the Gophers"
        picCoach.Print "from 1999-2001. Bickerstaff is the"
        picCoach.Print "youngest assistant coach in NBA"
        picCoach.Print "history."
    Case Is = "Ed Pickney"
        'This gives info for Pickney
        picCoach.Cls
        picCoach.Print "Pickney is an assistant coach for"
        picCoach.Print "the Minnesota Timberwolves. "
        picCoach.Print "Pickney attended Villanova where "
        picCoach.Print "he played basketball. Pickney is"
        picCoach.Print "the newest edition to the"
        picCoach.Print "Timberwolves coaching staff."
    Case Is = "Brent Haskins"
        'this gives info about Haskins
        picCoach.Cls
        picCoach.Print "Haskins is the head scout for the"
        picCoach.Print "Minnesota Timberwolves. "
        picCoach.Print "Previously Haskins was the "
        picCoach.Print "assistant coach for the Minnesota"
        picCoach.Print "Gophers. Haskins helps the "
        picCoach.Print "Timberwolves specifically during"
        picCoach.Print "summer league, draft time, and"
        picCoach.Print "the playoffs."
    Case Is = "Dave Vitel"
        'this gives info about Vitel
        picCoach.Cls
        picCoach.Print "Dave is the Timberwolves strength"
        picCoach.Print "and conditioning coach. Dave's"
        picCoach.Print "responsible for getting the team"
        picCoach.Print "jacked up. Despite what you may"
        picCoach.Print "think the players are not on the"
        picCoach.Print "juice, but owe their muscular"
        picCoach.Print "physique's to the help and "
        picCoach.Print "expertise of Dave."
    Case Is = "Greg Farnam"
        'this gives info about Farnam
        picCoach.Cls
        picCoach.Print "Farnam is the Timberwolves"
        picCoach.Print "athletic trainer. Farnam is"
        picCoach.Print "starting his ninth season"
        picCoach.Print "with the Timberwolves."
        picCoach.Print "Previously Farnam was the trainer"
        picCoach.Print "for the St.Paul Saints. Unfortunately"
        picCoach.Print "Greg attended that school which is"
        picCoach.Print "far inferior to SJU by the name of"
        picCoach.Print "St.Cloud State University."
    Case Else
        picCoach.Cls
        picCoach.Print "Check your Spelling!"
        
        
    
    End Select
End Sub

Private Sub cmdreturn_Click()
'This allows the user to return to the main page by leaving the coaching page
frmCoachingStaff.Visible = False
frmMainPage.Visible = True

End Sub
