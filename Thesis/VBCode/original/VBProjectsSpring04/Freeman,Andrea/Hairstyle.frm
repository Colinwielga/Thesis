VERSION 5.00
Begin VB.Form AndreaFreemanfrmHairstyle 
   BackColor       =   &H00C0C000&
   Caption         =   "Hair"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10875
   LinkTopic       =   "Form5"
   ScaleHeight     =   7905
   ScaleWidth      =   10875
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
      Left            =   7920
      Picture         =   "Hairstyle.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6600
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
      Left            =   7920
      Picture         =   "Hairstyle.frx":204A2
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5400
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
      Height          =   4335
      Left            =   2760
      ScaleHeight     =   4275
      ScaleWidth      =   3555
      TabIndex        =   6
      Top             =   3240
      Width           =   3615
   End
   Begin VB.OptionButton optcrazy 
      BackColor       =   &H00C0C000&
      Caption         =   "Crazy Hair"
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
      Left            =   480
      TabIndex        =   4
      Top             =   3000
      Width           =   1695
   End
   Begin VB.OptionButton optlong 
      BackColor       =   &H00C0C000&
      Caption         =   "Long Hair"
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
      Left            =   4920
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin VB.OptionButton optshort 
      BackColor       =   &H00C0C000&
      Caption         =   "Short Hair"
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
      Left            =   2520
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.OptionButton optdreadlocks 
      BackColor       =   &H00C0C000&
      Caption         =   "Dreadlocks"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7560
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.OptionButton optbald 
      BackColor       =   &H00C0C000&
      Caption         =   "Bald"
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
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Choose which picture best corresponds to your hairstyle:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   120
      Width           =   6855
   End
   Begin VB.Image Image5 
      Height          =   3000
      Left            =   120
      Picture         =   "Hairstyle.frx":39C6C
      Top             =   3480
      Width           =   2445
   End
   Begin VB.Image Image4 
      Height          =   1860
      Left            =   5040
      Picture         =   "Hairstyle.frx":51D0E
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Image Image3 
      Height          =   1575
      Left            =   2520
      Picture         =   "Hairstyle.frx":5AAC0
      Top             =   1320
      Width           =   1980
   End
   Begin VB.Image Image2 
      Height          =   3240
      Left            =   7080
      Picture         =   "Hairstyle.frx":64D6E
      Top             =   1080
      Width           =   2685
   End
   Begin VB.Image Image1 
      Height          =   1320
      Left            =   240
      Picture         =   "Hairstyle.frx":81550
      Top             =   1200
      Width           =   2010
   End
End
Attribute VB_Name = "AndreaFreemanfrmHairstyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjPersonalityAnalysis (Andrea Freeman's VB Project.vbp)
'Form Name: AndreaFreemanfrmHairstyle (Hairstyle.frm)
'Author: Andrea Freeman
'Date Written: March 10, 2004
'Purpose of Form: This form asks the user to choose their hairstyle
                  'and displays their choice along with the previous choices.
                  'It then provides the user with a button to proceed to
                  'the next form.

Private Sub cmdnext_Click()
'Hide the Hairstyle screen and show the
'Analysis screen.
AndreaFreemanfrmHairstyle.Hide
AndreaFreemanfrmAnalysis.Show

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

'Determine which Hairstyle the user has selected.
If optbald = True Then
    M = 1
End If
If optdreadlocks = True Then
    M = 2
End If
If optshort = True Then
    M = 3
End If
If optlong = True Then
    M = 4
End If
If optcrazy = True Then
    M = 5
End If

picResults.Print "Your favorite animal is a "; FavoriteAnimal(I); "."
picResults.Print "Your favorite color is "; FavoriteColor(J); "."
picResults.Print "Your dream vacation is "; DreamVacation(K); "."
picResults.Print "You are currently "; Mood(L); "."
picResults.Print "Your hairstyle is "; Hairstyle(M); "."

'Make the print button inaccessible and the next button accessible.
cmdprintselections.Enabled = False
cmdnext.Enabled = True

'Make the print button invisible and the next button visible.
cmdprintselections.Visible = False
cmdnext.Visible = True
End Sub

Private Sub Form_Load()
PATH = "M:\CS130\Freeman, Andrea\"

'Open "M:\CS130\VB Project\Hairstyle.txt" For Input As #1
'Open PATH & "Hairstyle.txt" For Input As #1
Open "N:\CS130\handin\Freeman, Andrea\Hairstyle.txt" For Input As #1

For M = 1 To 5
    Input #1, Hairstyle(M), HairstylePhrase(M) 'The information
        'about the hairstyle and it's corresponding phrase are now
        'available to be used.
Next M
Close #1

'Make the print and next buttons inaccessible.
cmdprintselections.Enabled = False
cmdnext.Enabled = False

cmdnext.Visible = False 'Make the next button invisible.
End Sub

Private Sub optbald_Click()
'Enable "Print Selection" button after a selection has been made.
cmdprintselections.Enabled = True
End Sub

Private Sub optcrazy_Click()
'Enable "Print Selection" button after a selection has been made.
cmdprintselections.Enabled = True
End Sub

Private Sub optdreadlocks_Click()
'Enable "Print Selection" button after a selection has been made.
cmdprintselections.Enabled = True
End Sub

Private Sub optlong_Click()
'Enable "Print Selection" button after a selection has been made.
cmdprintselections.Enabled = True
End Sub

Private Sub optshort_Click()
'Enable "Print Selection" button after a selection has been made.
cmdprintselections.Enabled = True
End Sub
