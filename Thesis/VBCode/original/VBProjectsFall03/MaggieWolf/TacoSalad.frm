VERSION 5.00
Begin VB.Form TacoSaladfm 
   BackColor       =   &H00800080&
   Caption         =   "Form2"
   ClientHeight    =   10605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13425
   LinkTopic       =   "Form2"
   ScaleHeight     =   10605
   ScaleWidth      =   13425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Quit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1200
      TabIndex        =   4
      Top             =   6360
      Width           =   2415
   End
   Begin VB.CommandButton ORecipes8 
      Caption         =   "What Other Recipes Can I Make?"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   600
      TabIndex        =   3
      Top             =   4560
      Width           =   3975
   End
   Begin VB.PictureBox Picture2 
      Height          =   2415
      Left            =   1320
      Picture         =   "TacoSalad.frx":0000
      ScaleHeight     =   2355
      ScaleWidth      =   2235
      TabIndex        =   2
      Top             =   2040
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   5175
      Left            =   5640
      Picture         =   "TacoSalad.frx":1BA7
      ScaleHeight     =   5115
      ScaleWidth      =   3795
      TabIndex        =   1
      Top             =   2160
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800080&
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
      Height          =   495
      Left            =   11040
      TabIndex        =   5
      Top             =   10080
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800080&
      Caption         =   "Taco Salad"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3000
      TabIndex        =   0
      Top             =   600
      Width           =   4695
   End
End
Attribute VB_Name = "TacoSaladfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Taco Salad (TacoSaladfm)
'Maggie Wolf
'The purpose of this form is to display the recipe.
'This button ends the program.
Private Sub ORecipes8_Click()
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
