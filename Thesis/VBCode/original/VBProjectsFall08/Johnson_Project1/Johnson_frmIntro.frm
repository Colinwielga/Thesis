VERSION 5.00
Object = "SoundRec"; "sndrec32.exe"
Begin VB.Form frmIntro 
   BackColor       =   &H80000009&
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   255
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   Picture         =   "Johnson_frmIntro.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdQuiz 
      Caption         =   "The Brewers Quiz"
      Height          =   1215
      Left            =   8280
      TabIndex        =   11
      Top             =   5280
      Width           =   2175
   End
   Begin VB.PictureBox Picture2 
      Height          =   3255
      Left            =   10560
      Picture         =   "Johnson_frmIntro.frx":25DB
      ScaleHeight     =   3195
      ScaleWidth      =   4275
      TabIndex        =   8
      Top             =   4440
      Width           =   4335
   End
   Begin VB.PictureBox Picture1 
      Height          =   3855
      Left            =   10680
      Picture         =   "Johnson_frmIntro.frx":5AC4
      ScaleHeight     =   3795
      ScaleWidth      =   4035
      TabIndex        =   7
      Top             =   360
      Width           =   4095
   End
   Begin VB.CommandButton cmdTicketPrice 
      Caption         =   "Ticket Pricing"
      Height          =   1095
      Left            =   8280
      TabIndex        =   6
      Top             =   3960
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit!"
      Height          =   1095
      Left            =   8280
      TabIndex        =   5
      Top             =   6720
      Width           =   2175
   End
   Begin VB.CommandButton cmdActiveRoster 
      Caption         =   "Active Roster"
      Height          =   975
      Left            =   8280
      TabIndex        =   4
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdPitchGo 
      Caption         =   "Brew Crew Pitching"
      Height          =   1095
      Left            =   8280
      TabIndex        =   3
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton cmdPlayer 
      BackColor       =   &H00FF0000&
      Caption         =   "Brewers Hitting Stats"
      Height          =   975
      Left            =   8280
      TabIndex        =   0
      Top             =   1440
      Width           =   2175
   End
   Begin SoundRecCtl.SoundRec SoundRec1 
      Height          =   480
      Left            =   10920
      OleObjectBlob   =   "Johnson_frmIntro.frx":D3BC
      TabIndex        =   12
      Top             =   8640
      Width           =   480
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      Caption         =   "Milwaukee Brewers Fan Club Program 2008 "
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1800
      TabIndex        =   10
      Top             =   9480
      Width           =   11655
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Project Made By: Matthew C. Johnson"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   10080
      TabIndex        =   9
      Top             =   7920
      Width           =   4815
   End
   Begin VB.Label HomePageLabel 
      BackColor       =   &H80000004&
      Caption         =   "http://milwaukee.brewers.mlb.com/index.jsp?c_id=mil"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   8640
      Width           =   7455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Go to the Homepage:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   7920
      Width           =   4335
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Milwaukee Brewers Fan Club Program 2008

'Form Name: Initial Page Form

'Author: Matthew Johnson

'Date Written: 10/31/2008

'Objective of the form: In this particular part of the program, I create a linking page that
'connects to many forms.  In this form, I also created a URL link so the user can access
'the Brewer homepage.

'Objective of the program: To demonstrate what I've learned in a fun and creative way.
'I'm a die hard brewers fan!

Option Explicit
'Here I declare a function that can make me turn a label into a URL link.
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Here I allow the user to access the Active Roster Form, and leave the initial page.
Private Sub cmdActiveRoster_Click()
    frmIntro.Hide
    frmRoster.Show
End Sub

Private Sub cmdCite_Click()

End Sub

'Here I allow the user to access the Brewer's Pitching Form, and leave the initial page
Private Sub cmdPitchGo_Click()
    frmIntro.Hide
    frmPitch.Show
End Sub
'Here I allow the user to access the Brewer' 2008 Hitting Statistics Form, and leave
'the initial page.
Private Sub cmdPlayer_Click()
    frmIntro.Hide
    frmPlayer.Show
End Sub
'Here I allow the user to quit the program.
Private Sub cmdQuit_Click()
End
End Sub
'Here I allow the user to access the Brewer's Trivia Form, and leave the initial page.
Private Sub cmdQuiz_Click()
    frmIntro.Hide
    frmQuiz.Show
End Sub
'Here I allow the user to access the Ticket Price calculator Form, and leave the initial page.
Private Sub cmdTicketPrice_Click()
    frmIntro.Hide
    frmTickets.Show
End Sub
'Here I create a URL link, so that the user can access the Brewer's official website.
Private Sub HomePageLabel_Click()
    With HomePageLabel
        Call ShellExecute(0&, vbNullString, .Caption, vbNullString, vbNullString, vbNormalFocus)
    End With
End Sub

'Here I give it a url label... After it's clicked, the font color of the caption becomes blue and the font becomes underlined.
Private Sub HomePageLabel_MouseUP(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With HomePageLabel
        .ForeColor = vbBlue
        .Font.Underline = True
    End With
End Sub
