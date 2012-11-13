VERSION 5.00
Begin VB.Form AndreaFreemanfrmFavoriteColor 
   BackColor       =   &H00C0C000&
   Caption         =   "Favorite Color"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10590
   LinkTopic       =   "Form2"
   ScaleHeight     =   8115
   ScaleWidth      =   10590
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
      Picture         =   "FavoriteColor.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
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
      Left            =   7920
      Picture         =   "FavoriteColor.frx":204A2
      Style           =   1  'Graphical
      TabIndex        =   12
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
      Height          =   3135
      Left            =   360
      ScaleHeight     =   3075
      ScaleWidth      =   6555
      TabIndex        =   11
      Top             =   4080
      Width           =   6615
   End
   Begin VB.OptionButton optyellow 
      BackColor       =   &H00C0C000&
      Caption         =   "Yellow"
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
      Left            =   8160
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
   End
   Begin VB.OptionButton optblue 
      BackColor       =   &H00C0C000&
      Caption         =   "Blue"
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
      Left            =   6360
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.OptionButton optblack 
      BackColor       =   &H00C0C000&
      Caption         =   "Black"
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
      Left            =   4440
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
   End
   Begin VB.OptionButton optred 
      BackColor       =   &H00C0C000&
      Caption         =   " Red"
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
      Left            =   2280
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.OptionButton optgreen 
      BackColor       =   &H00C0C000&
      Caption         =   " Green"
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
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H0000FFFF&
      Height          =   1695
      Left            =   8160
      ScaleHeight     =   1635
      ScaleWidth      =   1275
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00800000&
      Height          =   1695
      Left            =   6240
      ScaleHeight     =   1635
      ScaleWidth      =   1395
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      Height          =   1695
      Left            =   4200
      ScaleHeight     =   1635
      ScaleWidth      =   1395
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000000C0&
      Height          =   1695
      Left            =   2160
      ScaleHeight     =   1635
      ScaleWidth      =   1395
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000C000&
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1635
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Choose your favorite color:"
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
      Left            =   2520
      TabIndex        =   10
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "AndreaFreemanfrmFavoriteColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjPersonalityAnalysis (Andrea Freeman's VB Project.vbp)
'Form Name: AndreaFreemanfrmFavoriteColor (FavoriteColor.frm)
'Author: Andrea Freeman
'Date Written: March 9, 2004
'Purpose of Form: This form asks the user to choose their favorite color
                  'and displays their choice along with the previous one.
                  'It then provides the user with a button to proceed to
                  'the next form.


Private Sub cmdnext_Click()
'Hide the Favorite Color screen and show the
'Dream Vacation screen.
AndreaFreemanfrmFavoriteColor.Hide
AndreaFreemanfrmDreamVacation.Show

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

'Determine which Favorite Color the user has selected.
If optgreen = True Then
    J = 1
End If
If optred = True Then
    J = 2
End If
If optblack = True Then
    J = 3
End If
If optblue = True Then
    J = 4
End If
If optyellow = True Then
    J = 5
End If

picResults.Print "Your favorite animal is a "; FavoriteAnimal(I); "."
picResults.Print "Your favorite color is "; FavoriteColor(J); "."

'Make the print button inaccessible and the next button accessible.
cmdprintselections.Enabled = False
cmdnext.Enabled = True

'Make the print button invisible and the next button visible.
cmdprintselections.Visible = False
cmdnext.Visible = True
End Sub

Private Sub Form_Load()
PATH = "M:\CS130\Freeman, Andrea\"

'Open "M:\CS130\VB Project\FavoriteColor.txt" For Input As #1
'Open PATH & "FavoriteColor.txt" For Input As #1
Open "N:\CS130\handin\Freeman, Andrea\FavoriteColor.txt" For Input As #1

For J = 1 To 5
    Input #1, FavoriteColor(J), FavoriteColorPhrase(J) 'The information
        'about the favorite color and it's corresponding phrase are now
        'available to be used.
Next J
Close #1

'Make the print and next buttons inaccessible.
cmdprintselections.Enabled = False
cmdnext.Enabled = False

cmdnext.Visible = False 'Make the next button invisible.
End Sub

Private Sub optblack_Click()
'Enable "Print Selection" button after a selection has been made.
cmdprintselections.Enabled = True
End Sub

Private Sub optblue_Click()
'Enable "Print Selection" button after a selection has been made.
cmdprintselections.Enabled = True
End Sub

Private Sub optgreen_Click()
'Enable "Print Selection" button after a selection has been made.
cmdprintselections.Enabled = True
End Sub

Private Sub optred_Click()
'Enable "Print Selection" button after a selection has been made.
cmdprintselections.Enabled = True
End Sub

Private Sub optyellow_Click()
'Enable "Print Selection" button after a selection has been made.
cmdprintselections.Enabled = True
End Sub
