VERSION 5.00
Begin VB.Form ChickenHotdishfm 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form1"
   ClientHeight    =   10575
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   ScaleHeight     =   10575
   ScaleWidth      =   12060
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   4935
      Left            =   6120
      Picture         =   "ChickenHotdish.frx":0000
      ScaleHeight     =   4875
      ScaleWidth      =   3675
      TabIndex        =   4
      Top             =   2040
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Tiempo"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1200
      TabIndex        =   3
      Top             =   6720
      Width           =   2295
   End
   Begin VB.CommandButton ORecipes2 
      Caption         =   "What other recipes are there to make?"
      BeginProperty Font 
         Name            =   "Tiempo"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   720
      TabIndex        =   1
      Top             =   4320
      Width           =   3735
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   1560
      Picture         =   "ChickenHotdish.frx":1DCC
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Maggie Wolf"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   5
      Top             =   10200
      Width           =   1455
   End
   Begin VB.Label Chickendish 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Chicken Hotdish"
      BeginProperty Font 
         Name            =   "Tiempo"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   2
      Top             =   480
      Width           =   5415
   End
End
Attribute VB_Name = "ChickenHotdishfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Chicken Hotdish (ChickenHotdishfm)
'Maggie Wolf
'The purpose of this form is to display the recipe.
'This button ends the program.
Private Sub Command1_Click()
    End
End Sub
'This button displays the form that holds the other recipe choices.
Private Sub ORecipes2_Click()
    Firstfm.Hide
    RecipeChoicesfm.Show
    ChickenHotdishfm.Hide
    Lasagnafm.Hide
    TacoSaladfm.Hide
    Fajitasfm.Hide
    Mushfm.Hide
    SloppyJoesfm.Hide
    ChickenStuffingfm.Hide
    StirFryfm.Hide
End Sub
