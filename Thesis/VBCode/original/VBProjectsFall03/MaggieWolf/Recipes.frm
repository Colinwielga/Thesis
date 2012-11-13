VERSION 5.00
Begin VB.Form Lasagnafm 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form1"
   ClientHeight    =   10920
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12960
   LinkTopic       =   "Form1"
   ScaleHeight     =   10920
   ScaleWidth      =   12960
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   6735
      Left            =   5640
      Picture         =   "Recipes.frx":0000
      ScaleHeight     =   6675
      ScaleWidth      =   5235
      TabIndex        =   4
      Top             =   1680
      Width           =   5295
   End
   Begin VB.CommandButton Quit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1560
      TabIndex        =   3
      Top             =   6480
      Width           =   2775
   End
   Begin VB.CommandButton ORecipes1 
      Caption         =   "What other recipes are there to make?"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   840
      TabIndex        =   2
      Top             =   3840
      Width           =   3975
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   2040
      Picture         =   "Recipes.frx":2F26
      ScaleHeight     =   1515
      ScaleWidth      =   1395
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
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
      Left            =   10560
      TabIndex        =   5
      Top             =   10440
      Width           =   1575
   End
   Begin VB.Label Lasagna 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Lasagna"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4080
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "Lasagnafm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Lasagna (Lasagnafm)
'Maggie Wolf
'The purpose of this form is to display the recipe.
'This button ends the program.
Private Sub ORecipes1_Click()
'This button displays the form that holds the other recipe choices.
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

Private Sub Quit_Click()
    End
End Sub
