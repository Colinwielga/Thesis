VERSION 5.00
Begin VB.Form frmSeasonsOverview
   Caption         =   "Form1"
   ClientHeight    =   9720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   Picture         =   "frmSeasonsOverview.frx":0000
   ScaleHeight     =   9720
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit
      Caption         =   "Quit"
      Height          =   495
      Left            =   3480
      TabIndex        =   9
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdReturn
      Caption         =   "Return to Main Menu"
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Top             =   1200
      Width           =   1455
   End
   Begin VB.PictureBox picResults
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   240
      ScaleHeight     =   3555
      ScaleWidth      =   8235
      TabIndex        =   7
      Top             =   5520
      Width           =   8295
   End
   Begin VB.CommandButton cmdSeason2
      Caption         =   "Season 2"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdSeason7
      Caption         =   "Season 7"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdSeason6
      Caption         =   "Season 6"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdSeason5
      Caption         =   "Season 5"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdSeason4
      Caption         =   "Season 4"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdSeason3
      Caption         =   "Season 3"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdSeason1
      Caption         =   "Season 1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmSeasonsOverview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdReturn_Click()
    frmSeasonsOverview.Hide
    frmMain.Show
End Sub

Private Sub cmdSeason1_Click()
    picResults.Cls
    picResults.Print "Season 1 starts out with Vincent Chase and his 'entourage'"
    picResults.Print "going through the Hollywood lifestyle of being a movie star"
    picResults.Print "and the friends of a famous movie star."
    picResults.Print "After the open of Head On, Vince decides to do a passion"
    picResults.Print "project by making Queens Boulevard with up coming director"
    picResults.Print "Billy Walsh."
    picResults.Print "Eric works as Vince's manager but wants the title."
    picResults.Print "Meanwhile the rest of the gang is perfectly happy just mooching"
    picResults.Print "off of Vince's accomplishments and fame."
    picResults.Print "The first season is loosely based off of Mark Wahlberg's career."
End Sub

Private Sub cmdSeason3_Click()
    picResults.Cls
    picResults.Print "Ari has finally started his own agency, even though it's small"
    picResults.Print "Vince has chosen to stay with him. The Aquaman premiere has"
    picResults.Print "arrived and after it shatters the record book for highest"
    picResults.Print "opening of all time, Vince decides he wants to do another"
    picResults.Print "passion project Medillin. But with the success of Aquaman"
    picResults.Print "the studio wants to get started on Aquaman 2 giving Vince no"
    picResults.Print "no time to do Medillin. He gets fired from A2 and then loses"
    picResults.Print "Medillin as well. After Ari loses Vince's next possible project"
    picResults.Print "Ari is fired and Amanda is hired. Drama gets a hit TV series"
End Sub

Private Sub cmdSeason4_Click()
picResults.Cls
picResults.Print "The season begins with Vince shooting Medellin down in Columbia."
picResults.Print "The crew does a lot of relaxing and partying when they get back stateside."
picResults.Print "Eric also dives further into his managing business."
picResults.Print "Nothing major happens in regards to the plot until they go to Canes."
picResults.Print "The moment they have all been waiting for: Medellin is premierd."
picResults.Print "Medellin is a bomb at the Canes Film Festival."
picResults.Print "Drama finds love in Canes and it continues into the 5th season."
picResults.Print "The season ends as they are offered $1 for the movie."
End Sub

Private Sub cmdSeason2_Click()
    picResults.Cls
    picResults.Print "With the buzz of Queens Boulevard at Sundance Film Festival"
    picResults.Print "Vince is the talk of the town and lands a James Cameron film"
    picResults.Print "playing as Aquaman. With the new movie deal, Vince treats"
    picResults.Print "his friends to a new house and new toys. An old love in"
    picResults.Print "Mandy Moore returns to the picture only to leave again"
    picResults.Print "while Ari gets fired from his agency."
End Sub

Private Sub cmdSeason7_Click()
    picResults.Cls
    picResults.Print "Season 7 is set to air in the Summer of 2010."
End Sub
