VERSION 5.00
Begin VB.Form AndreaFreemanfrmFavoriteAnimal 
   BackColor       =   &H00C0C000&
   Caption         =   "Favorite Animal"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   10575
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
      Picture         =   "FavoriteAnimal.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6360
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
      Height          =   2655
      Left            =   360
      ScaleHeight     =   2595
      ScaleWidth      =   5955
      TabIndex        =   6
      Top             =   4440
      Width           =   6015
   End
   Begin VB.CommandButton cmdprintselections 
      BackColor       =   &H0000C000&
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
      Picture         =   "FavoriteAnimal.frx":204A2
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4920
      Width           =   2055
   End
   Begin VB.OptionButton optcamel 
      BackColor       =   &H00C0C000&
      Caption         =   "Camel"
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
      Left            =   8280
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
   End
   Begin VB.OptionButton optdog 
      BackColor       =   &H00C0C000&
      Caption         =   "Dog"
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
      Left            =   6120
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.OptionButton optcat 
      BackColor       =   &H00C0C000&
      Caption         =   " Cat"
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
      Top             =   1440
      Width           =   975
   End
   Begin VB.OptionButton optparrot 
      BackColor       =   &H00C0C000&
      Caption         =   "Parrot"
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
      Left            =   2040
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.OptionButton optmonkey 
      BackColor       =   &H00C0C000&
      Caption         =   "Monkey"
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
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Choose your favorite animal:"
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
      Left            =   2640
      TabIndex        =   8
      Top             =   240
      Width           =   3615
   End
   Begin VB.Image Image5 
      Height          =   1605
      Left            =   8160
      Picture         =   "FavoriteAnimal.frx":39C6C
      Top             =   2280
      Width           =   1785
   End
   Begin VB.Image Image4 
      Height          =   1875
      Left            =   6120
      Picture         =   "FavoriteAnimal.frx":43326
      Top             =   2040
      Width           =   1425
   End
   Begin VB.Image Image3 
      Height          =   1680
      Left            =   3600
      Picture         =   "FavoriteAnimal.frx":4C008
      Top             =   2160
      Width           =   2235
   End
   Begin VB.Image Image2 
      Height          =   1905
      Left            =   1800
      Picture         =   "FavoriteAnimal.frx":5844A
      Top             =   1920
      Width           =   1395
   End
   Begin VB.Image Image1 
      Height          =   2235
      Left            =   120
      Picture         =   "FavoriteAnimal.frx":60F74
      Top             =   1680
      Width           =   1485
   End
End
Attribute VB_Name = "AndreaFreemanfrmFavoriteAnimal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjPersonalityAnalysis (Andrea Freeman's VB Project.vbp)
'Form Name: AndreaFreemanfrmFavoriteAnimal (FavoriteAnimal.frm)
'Author: Andrea Freeman
'Date Written: March 9, 2004
'Purpose of Form: This form asks the user to choose their favorite animal
                  'and displays their choice. It then provides the user with
                  'a button to proceed to the next form.

Private Sub cmdprintselections_Click()

'Clear whatever may be in picResults for repeated use.
picResults.Cls

'Determine which Favorite Animal the user has selected.
If optmonkey = True Then
    I = 1
End If
If optparrot = True Then
    I = 2
End If
If optcat = True Then
    I = 3
End If
If optdog = True Then
    I = 4
End If
If optcamel = True Then
    I = 5
End If

picResults.Print "Your favorite animal is a "; FavoriteAnimal(I); "."

'Make the print button inaccessible and the next button accessible.
cmdprintselections.Enabled = False
cmdnext.Enabled = True

'Make the print button invisible and the next button visible.
cmdprintselections.Visible = False
cmdnext.Visible = True
End Sub


Private Sub cmdnext_Click()
'Hide the Favorite Animal screen and show the
'Favorite Color screen.
AndreaFreemanfrmFavoriteAnimal.Hide
AndreaFreemanfrmFavoriteColor.Show

'Make the print button accessible and visible for repeated use.
cmdprintselections.Enabled = True
cmdprintselections.Visible = True

'Reset the next button as inaccessible and invisible for repeated use.
cmdnext.Enabled = False
cmdnext.Visible = False
End Sub

Private Sub Form_Load()
PATH = "M:\CS130\Freeman, Andrea\"

'Open "M:\CS130\VB Project\FavoriteAnimal.txt" For Input As #1
'Open PATH & "FavoriteAnimal.txt" For Input As #1
Open "N:\CS130\handin\Freeman, Andrea\FavoriteAnimal.txt" For Input As #1

For I = 1 To 5
    Input #1, FavoriteAnimal(I), FavoriteAnimalPhrase(I) 'The information
        'about the favorite animal and it's corresponding phrase are now
        'available to be used.
Next I
Close #1

'Make the print and next buttons inaccessible.
cmdprintselections.Enabled = False
cmdnext.Enabled = False

'Reset the next button as inaccessible and invisible for repeated use.
cmdnext.Enabled = False
cmdnext.Visible = False
End Sub


Private Sub optcamel_Click()
'Enable "Print Selection" button after a selection has been made.
cmdprintselections.Enabled = True
End Sub

Private Sub optcat_Click()
'Enable "Print Selection" button after a selection has been made.
cmdprintselections.Enabled = True
End Sub

Private Sub optdog_Click()
'Enable "Print Selection" button after a selection has been made.
cmdprintselections.Enabled = True
End Sub

Private Sub optmonkey_Click()
'Enable "Print Selection" button after a selection has been made.
cmdprintselections.Enabled = True
End Sub

Private Sub optparrot_Click()
'Enable "Print Selection" button after a selection has been made.
cmdprintselections.Enabled = True
End Sub
