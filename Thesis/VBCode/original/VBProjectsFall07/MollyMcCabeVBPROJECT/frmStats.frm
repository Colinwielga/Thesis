VERSION 5.00
Begin VB.Form frmStats 
   BackColor       =   &H00400000&
   Caption         =   "Player Statistics"
   ClientHeight    =   4035
   ClientLeft      =   4110
   ClientTop       =   4485
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   8325
   Begin VB.CommandButton cmdStats 
      BackColor       =   &H000000C0&
      Caption         =   "See Stats"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000C0&
      Caption         =   "Exit Twins Territory"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack2 
      BackColor       =   &H000000C0&
      Caption         =   "Back to Twins Territory"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000C0&
      Caption         =   "Choose Another Twin"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3000
      Width           =   1575
   End
   Begin VB.PictureBox picDebut 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      ScaleHeight     =   315
      ScaleWidth      =   1875
      TabIndex        =   15
      Top             =   2520
      Width           =   1935
   End
   Begin VB.PictureBox picRightLeft 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      ScaleHeight     =   315
      ScaleWidth      =   1875
      TabIndex        =   14
      Top             =   2520
      Width           =   1935
   End
   Begin VB.PictureBox picWeight 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      ScaleHeight     =   315
      ScaleWidth      =   1875
      TabIndex        =   13
      Top             =   1800
      Width           =   1935
   End
   Begin VB.PictureBox picHeight 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      ScaleHeight     =   315
      ScaleWidth      =   1875
      TabIndex        =   12
      Top             =   1800
      Width           =   1935
   End
   Begin VB.PictureBox picTown 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      ScaleHeight     =   315
      ScaleWidth      =   2955
      TabIndex        =   11
      Top             =   1080
      Width           =   3015
   End
   Begin VB.PictureBox picFullName 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      ScaleHeight     =   315
      ScaleWidth      =   2475
      TabIndex        =   10
      Top             =   360
      Width           =   2535
   End
   Begin VB.PictureBox picBirthday 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      ScaleHeight     =   315
      ScaleWidth      =   1875
      TabIndex        =   9
      Top             =   1080
      Width           =   1935
   End
   Begin VB.PictureBox picName 
      BackColor       =   &H000000C0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      ScaleHeight     =   435
      ScaleWidth      =   1995
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.PictureBox picPhoto 
      Height          =   1935
      Left            =   720
      ScaleHeight     =   1875
      ScaleWidth      =   1275
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblDebut 
      BackColor       =   &H000000C0&
      Caption         =   "MLB Debut:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   8
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblRightLeft 
      BackColor       =   &H000000C0&
      Caption         =   "Bats/Throws:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblWeight 
      BackColor       =   &H000000C0&
      Caption         =   "Weight:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   6
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblHeight 
      BackColor       =   &H000000C0&
      Caption         =   "Height:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblTown 
      BackColor       =   &H000000C0&
      Caption         =   "Birthplace:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblBirth 
      BackColor       =   &H000000C0&
      Caption         =   "Date of Birth:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblFullName 
      BackColor       =   &H000000C0&
      Caption         =   "Full Name:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
'clear photo and stat boxes
picName.Cls
picPhoto.Picture = Nothing
picFullName.Cls
picBirthday.Cls
picTown.Cls
picHeight.Cls
picWeight.Cls
picRightLeft.Cls
picDebut.Cls
'goes back to FavTwin form
    frmStats.Hide
    frmFavTwin.Show
End Sub

Private Sub cmdBack2_Click()
'goes back to Main form
    frmStats.Hide
    frmMain.Show
End Sub

Private Sub cmdExit_Click()
'ends program
    End
End Sub

Private Sub cmdStats_Click()
'print photo and stats on specific player "I"
picName.Print player(I)
picPhoto.Picture = LoadPicture(App.Path & "\" & photo(I))
picFullName.Print fullname(I)
picBirthday.Print birthday(I)
picTown.Print town(I)
picHeight.Print ft(I)
picWeight.Print pounds(I)
picRightLeft.Print rightleft(I)
picDebut.Print debut(I)
End Sub
