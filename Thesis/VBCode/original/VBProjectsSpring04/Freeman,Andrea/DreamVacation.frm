VERSION 5.00
Begin VB.Form AndreaFreemanfrmDreamVacation 
   BackColor       =   &H00C0C000&
   Caption         =   "Dream Vacation"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10740
   LinkTopic       =   "Form3"
   ScaleHeight     =   7755
   ScaleWidth      =   10740
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
      Picture         =   "DreamVacation.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6480
      Width           =   2175
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
      Picture         =   "DreamVacation.frx":204A2
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5160
      Width           =   2175
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
      Height          =   2775
      Left            =   480
      ScaleHeight     =   2715
      ScaleWidth      =   6675
      TabIndex        =   6
      Top             =   4560
      Width           =   6735
   End
   Begin VB.OptionButton optkenya 
      BackColor       =   &H00C0C000&
      Caption         =   "Kenya"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.OptionButton optmexico 
      BackColor       =   &H00C0C000&
      Caption         =   "Mexico"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   3
      Top             =   3120
      Width           =   1095
   End
   Begin VB.OptionButton optitaly 
      BackColor       =   &H00C0C000&
      Caption         =   "Italy"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   1440
      Width           =   855
   End
   Begin VB.OptionButton optalaska 
      BackColor       =   &H00C0C000&
      Caption         =   "Alaska"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   3120
      Width           =   1095
   End
   Begin VB.OptionButton opttahiti 
      BackColor       =   &H00C0C000&
      Caption         =   "Tahiti"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Choose your dream vacation:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   5
      Top             =   240
      Width           =   3615
   End
   Begin VB.Image Image6 
      Height          =   1590
      Left            =   7560
      Picture         =   "DreamVacation.frx":39C6C
      Top             =   1200
      Width           =   2115
   End
   Begin VB.Image Image4 
      Height          =   1245
      Left            =   4800
      Picture         =   "DreamVacation.frx":44C3E
      Top             =   3000
      Width           =   1845
   End
   Begin VB.Image Image3 
      Height          =   1485
      Left            =   4080
      Picture         =   "DreamVacation.frx":4C51C
      Top             =   1200
      Width           =   2235
   End
   Begin VB.Image Image2 
      Height          =   1410
      Left            =   1200
      Picture         =   "DreamVacation.frx":5729E
      Top             =   2880
      Width           =   2115
   End
   Begin VB.Image Image1 
      Height          =   1155
      Left            =   1200
      Picture         =   "DreamVacation.frx":60E90
      Top             =   1200
      Width           =   1725
   End
End
Attribute VB_Name = "AndreaFreemanfrmDreamVacation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjPersonalityAnalysis (Andrea Freeman's VB Project.vbp)
'Form Name: AndreaFreemanfrmDreamVacation (DreamVacation.frm)
'Author: Andrea Freeman
'Date Written: March 9, 2004
'Purpose of Form: This form asks the user to choose their dream vacation
                  'and displays their choice along with the previous choices.
                  'It then provides the user with a button to proceed to
                  'the next form.
                  
Private Sub cmdnext_Click()
'Hide the Dream Vacation screen and show the
'Mood screen.
AndreaFreemanfrmDreamVacation.Hide
AndreaFreemanfrmMood.Show

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

'Determine which Dream Vacation the user has selected.
If opttahiti = True Then
    K = 1
End If
If optalaska = True Then
    K = 2
End If
If optitaly = True Then
    K = 3
End If
If optmexico = True Then
    K = 4
End If
If optkenya = True Then
    K = 5
End If

picResults.Print "Your favorite animal is a "; FavoriteAnimal(I); "."
picResults.Print "Your favorite color is "; FavoriteColor(J); "."
picResults.Print "Your dream vacation is "; DreamVacation(K); "."

'Make the print button inaccessible and the next button accessible.
cmdprintselections.Enabled = False
cmdnext.Enabled = True

'Make the print button invisible and the next button visible.
cmdprintselections.Visible = False
cmdnext.Visible = True
End Sub


Private Sub Form_Load()
PATH = "M:\CS130\Freeman, Andrea\"

'Open "M:\CS130\VB Project\DreamVacation.txt" For Input As #1
'Open PATH & "DreamVacation.txt" For Input As #1
Open "N:\CS130\handin\Freeman, Andrea\DreamVacation.txt" For Input As #1

For K = 1 To 5
    Input #1, DreamVacation(K), DreamVacationPhrase(K) 'The information
        'about the dream vacation and it's corresponding phrase are now
        'available to be used.
Next K
Close #1

'Make the print and next buttons inaccessible.
cmdprintselections.Enabled = False
cmdnext.Enabled = False

cmdnext.Visible = False 'Make the next button invisible.
End Sub

Private Sub optalaska_Click()
'Enable "Print Selection" button after a selection has been made.
cmdprintselections.Enabled = True
End Sub

Private Sub optitaly_Click()
'Enable "Print Selection" button after a selection has been made.
cmdprintselections.Enabled = True
End Sub

Private Sub optkenya_Click()
'Enable "Print Selection" button after a selection has been made.
cmdprintselections.Enabled = True
End Sub

Private Sub optmexico_Click()
'Enable "Print Selection" button after a selection has been made.
cmdprintselections.Enabled = True
End Sub

Private Sub opttahiti_Click()
'Enable "Print Selection" button after a selection has been made.
cmdprintselections.Enabled = True
End Sub
