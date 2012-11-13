VERSION 5.00
Begin VB.Form AndreaFreemanfrmMood 
   BackColor       =   &H00C0C000&
   Caption         =   "Mood"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10530
   LinkTopic       =   "Form4"
   ScaleHeight     =   7710
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdnext 
      Caption         =   "Next Category"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8040
      Picture         =   "Mood.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6360
      Width           =   2055
   End
   Begin VB.CommandButton cmdprintselections 
      Caption         =   "Print Your Selections"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8040
      Picture         =   "Mood.frx":204A2
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5160
      Width           =   2055
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
      Height          =   3135
      Left            =   480
      ScaleHeight     =   3075
      ScaleWidth      =   5955
      TabIndex        =   6
      Top             =   3840
      Width           =   6015
   End
   Begin VB.OptionButton optangry 
      BackColor       =   &H00C0C000&
      Caption         =   "Angry"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.OptionButton optsad 
      BackColor       =   &H00C0C000&
      Caption         =   "Sad"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.OptionButton optneutral 
      BackColor       =   &H00C0C000&
      Caption         =   "Neutral"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.OptionButton opthappy 
      BackColor       =   &H00C0C000&
      Caption         =   "Happy"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.OptionButton optecstatic 
      BackColor       =   &H00C0C000&
      Caption         =   "Ecstatic"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Choose which picture best corresponds to your current mood:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   7815
   End
   Begin VB.Image Image5 
      Height          =   1260
      Left            =   7560
      Picture         =   "Mood.frx":39C6C
      Top             =   1560
      Width           =   1680
   End
   Begin VB.Image Image4 
      Height          =   2010
      Left            =   5640
      Picture         =   "Mood.frx":40AEE
      Top             =   1560
      Width           =   1560
   End
   Begin VB.Image Image3 
      Height          =   1635
      Left            =   4320
      Picture         =   "Mood.frx":4AE80
      Top             =   1560
      Width           =   870
   End
   Begin VB.Image Image2 
      Height          =   1890
      Left            =   1680
      Picture         =   "Mood.frx":4F9B2
      Top             =   1560
      Width           =   2235
   End
   Begin VB.Image Image1 
      Height          =   1620
      Left            =   120
      Picture         =   "Mood.frx":5D674
      Top             =   1560
      Width           =   1170
   End
End
Attribute VB_Name = "AndreaFreemanfrmMood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjPersonalityAnalysis (Andrea Freeman's VB Project.vbp)
'Form Name: AndreaFreemanfrmMood (Mood.frm)
'Author: Andrea Freeman
'Date Written: March 9, 2004
'Purpose of Form: This form asks the user to choose their current mood
                  'and displays their choice along with the previous choices.
                  'It then provides the user with a button to proceed to
                  'the next form.


Private Sub cmdnext_Click()
'Hide the Mood screen and show the
'Hairstyle screen.
AndreaFreemanfrmMood.Hide
AndreaFreemanfrmHairstyle.Show

'Make the print button accessible and visible for repeated use.
cmdprintselections.Enabled = True
cmdprintselections.Visible = True

'Reset the next button as inaccessible and invisible for repeated use.
cmdnext.Enabled = False
cmdnext.Visible = False
End Sub

Private Sub cmdprintselections_Click()
'Clear whatever may be in picResults for repeated use.
picResults.Cls

'Determine which Mood the user has selected.
If optecstatic = True Then
    L = 1
End If
If opthappy = True Then
    L = 2
End If
If optneutral = True Then
    L = 3
End If
If optsad = True Then
    L = 4
End If
If optangry = True Then
    L = 5
End If

picResults.Print "Your favorite animal is a "; FavoriteAnimal(I); "."
picResults.Print "Your favorite color is "; FavoriteColor(J); "."
picResults.Print "Your dream vacation is "; DreamVacation(K); "."
picResults.Print "You are currently "; Mood(L); "."

'Make the print button inaccessible and the next button accessible.
cmdprintselections.Enabled = False
cmdnext.Enabled = True

'Make the print button invisible and the next button visible.
cmdprintselections.Visible = False
cmdnext.Visible = True
End Sub

Private Sub Form_Load()
PATH = "M:\CS130\Freeman, Andrea\"

'Open "M:\CS130\VB Project\Mood.txt" For Input As #1
'Open PATH & "Mood.txt" For Input As #1
Open "N:\CS130\handin\Freeman, Andrea\Mood.txt" For Input As #1

For L = 1 To 5
    Input #1, Mood(L), MoodPhrase(L) 'The information
        'about mood and it's corresponding phrase are now
        'available to be used.
Next L
Close #1

'Make the print and next buttons inaccessible.
cmdprintselections.Enabled = False
cmdnext.Enabled = False

cmdnext.Visible = False 'Make the next button invisible.
End Sub

Private Sub optangry_Click()
'Enable "Print Selection" button after a selection has been made.
cmdprintselections.Enabled = True
End Sub

Private Sub optecstatic_Click()
'Enable "Print Selection" button after a selection has been made.
cmdprintselections.Enabled = True
End Sub

Private Sub opthappy_Click()
'Enable "Print Selection" button after a selection has been made.
cmdprintselections.Enabled = True
End Sub

Private Sub optneutral_Click()
'Enable "Print Selection" button after a selection has been made.
cmdprintselections.Enabled = True
End Sub

Private Sub optsad_Click()
'Enable "Print Selection" button after a selection has been made.
cmdprintselections.Enabled = True
End Sub
