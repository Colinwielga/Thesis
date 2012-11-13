VERSION 5.00
Begin VB.Form Mushfm 
   BackColor       =   &H00FF8080&
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13860
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   13860
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   4095
      Left            =   6120
      Picture         =   "Mush.frx":0000
      ScaleHeight     =   4035
      ScaleWidth      =   4155
      TabIndex        =   4
      Top             =   3240
      Width           =   4215
   End
   Begin VB.CommandButton Quit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "@Gungsuh"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   960
      TabIndex        =   3
      Top             =   6600
      Width           =   2895
   End
   Begin VB.CommandButton ORecipes4 
      Caption         =   "What other recipes can I make?"
      BeginProperty Font 
         Name            =   "@Gungsuh"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   4560
      Width           =   4335
   End
   Begin VB.PictureBox Picture2 
      Height          =   1815
      Left            =   1800
      Picture         =   "Mush.frx":18B0
      ScaleHeight     =   1755
      ScaleWidth      =   1155
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
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
      Left            =   11160
      TabIndex        =   5
      Top             =   10680
      Width           =   1455
   End
   Begin VB.Label Mush 
      BackColor       =   &H00FF8080&
      Caption         =   "Mush"
      BeginProperty Font 
         Name            =   "@Gungsuh"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3480
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
End
Attribute VB_Name = "Mushfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mush (Mushfm)
'Maggie Wolf
'The purpose of this form is to display the recipe.
'This button ends the program.
Private Sub ORecipes4_Click()
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
