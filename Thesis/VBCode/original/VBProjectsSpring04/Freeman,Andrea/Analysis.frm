VERSION 5.00
Begin VB.Form AndreaFreemanfrmAnalysis 
   BackColor       =   &H00C0C000&
   Caption         =   "Analysis"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10980
   LinkTopic       =   "Form6"
   ScaleHeight     =   8370
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdviewpics 
      Caption         =   "See all the pictures that you chose!"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8760
      Picture         =   "Analysis.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton cmddisplay 
      Caption         =   "Display Your Personality Analysis"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8760
      Picture         =   "Analysis.frx":C699
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FF80FF&
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   240
      ScaleHeight     =   7875
      ScaleWidth      =   7995
      TabIndex        =   3
      Top             =   240
      Width           =   8055
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8760
      Picture         =   "Analysis.frx":13E6D
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton cmdstart 
      Caption         =   "Start Over"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8760
      Picture         =   "Analysis.frx":1CB5A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton cmdlucky 
      Caption         =   "Compute Your Lucky Number!"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8760
      Picture         =   "Analysis.frx":1F7C1
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Width           =   1815
   End
End
Attribute VB_Name = "AndreaFreemanfrmAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjPersonalityAnalysis (Andrea Freeman's VB Project.vbp)
'Form Name: AndreaFreemanfrmAnalysis (Analysis.frm)
'Author: Andrea Freeman
'Date Written: March 11, 2004
'Purpose of Form: This form displays the personality analysis of the user by
                  'using the choices they made on the previous forms.
                  'It also calculates the user's lucky number and provides
                  'them with a button to either begin the program again or quit.
                  
Private Sub cmddisplay_Click()
'Display the personality analysis of the user.
picResults.Print "Your Personality Analysis:"
picResults.Print "____________________________________________________________"
picResults.Print FavoriteAnimalPhrase(I)
picResults.Print FavoriteColorPhrase(J)
picResults.Print DreamVacationPhrase(K)
picResults.Print MoodPhrase(L)
picResults.Print HairstylePhrase(M)

'Make the lucky number button accessible and visible.
cmdlucky.Enabled = True
cmdlucky.Visible = True

'Make the display button inaccessible and invisible.
cmddisplay.Visible = False
cmddisplay.Enabled = False
End Sub

Private Sub cmdlucky_Click()
'Declare variables that are associated just with this button.
Dim Day As Integer
Dim Month As Integer
Dim Year As Integer
Dim LuckyNumber As Integer

Day = InputBox("Please Enter Your Day of Birth", "Birthday")
Month = InputBox("Please Enter Your Month of Birth Numerically", "Birthmonth")
Year = InputBox("Please Enter the last 2 Digits of Your Birth Year", "Birthyear")

LuckyNumber = Day * Month / Year * 45

picResults.Print "____________________________________________________________"
picResults.Print "Your Lucky Number is "; LuckyNumber; "."

'Make the lucky number button inaccessible and invisible.
cmdlucky.Visible = False
cmdlucky.Enabled = False

'Make the start over and quit buttons accessible and visible.
cmdviewpics.Enabled = True
cmdviewpics.Visible = True
End Sub


Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdstart_Click()
picResults.Cls
'Hide the Analysis screen and show
'the Favorite Animal screen if the user wishes
'to use the program again.
AndreaFreemanfrmAnalysis.Hide
AndreaFreemanfrmFavoriteAnimal.Show

'Make the display button accessible and visible for repeated use.
cmddisplay.Enabled = True
cmddisplay.Visible = True

'Make the start and quit buttons inaccessible and invisible for repeated use.
cmdstart.Enabled = False
cmdstart.Visible = False
cmdquit.Enabled = False
cmdquit.Visible = False


End Sub

Private Sub cmdviewpics_Click()
'Make the Analysis form invisible and the Pictures form visible.
AndreaFreemanfrmAnalysis.Visible = False
AndreaFreemanfrmpictures.Visible = True

'Disable the View pictures button and make it invisible.
cmdviewpics.Enabled = False
cmdviewpics.Visible = False

'Make the start and quit buttons accessible and visible.
cmdstart.Enabled = True
cmdquit.Enabled = True
cmdstart.Visible = True
cmdquit.Visible = True

End Sub

Private Sub Form_Load()
'Make the lucky number, start over, and quit, and view picture buttons inaccessible.
cmdlucky.Enabled = False
cmdstart.Enabled = False
cmdquit.Enabled = False
cmdviewpics.Enabled = False

'Make the lucky number, start over, quit, and view picture buttons invisible.
cmdlucky.Visible = False
cmdstart.Visible = False
cmdquit.Visible = False
cmdviewpics.Visible = False
End Sub
