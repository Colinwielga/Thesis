VERSION 5.00
Begin VB.Form RecipeChoicesfm 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14250
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   14250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Quit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4680
      TabIndex        =   10
      Top             =   6600
      Width           =   2895
   End
   Begin VB.CommandButton firstgo 
      Caption         =   "Back to the first screen"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4680
      TabIndex        =   9
      Top             =   3480
      Width           =   2895
   End
   Begin VB.CommandButton SloppyJoesgo 
      Caption         =   "Sloppy Joes"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8640
      TabIndex        =   8
      Top             =   7440
      Width           =   3855
   End
   Begin VB.CommandButton TacoSaladgo 
      Caption         =   "Taco Salad"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8640
      TabIndex        =   7
      Top             =   5880
      Width           =   3855
   End
   Begin VB.CommandButton Fajitasgo 
      Caption         =   "Fajitas"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8640
      TabIndex        =   6
      Top             =   4320
      Width           =   3855
   End
   Begin VB.CommandButton ChickenHotdishgo 
      Caption         =   "Chicken Hotdish"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8640
      TabIndex        =   5
      Top             =   2520
      Width           =   3855
   End
   Begin VB.CommandButton lasagnago 
      Caption         =   "Lasagna"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   4
      Top             =   7440
      Width           =   3495
   End
   Begin VB.CommandButton StirFrygo 
      Caption         =   "Stir Fry"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   3
      Top             =   5880
      Width           =   3495
   End
   Begin VB.CommandButton chickenstuffinggo 
      Caption         =   "Chicken and Stuffing"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   2
      Top             =   4320
      Width           =   3495
   End
   Begin VB.CommandButton mushgo 
      Caption         =   "Mush"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   1
      Top             =   2640
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
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
      Left            =   12000
      TabIndex        =   11
      Top             =   10920
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Recipe Choices"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3000
      TabIndex        =   0
      Top             =   480
      Width           =   6855
   End
End
Attribute VB_Name = "RecipeChoicesfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Recipe Choices (RecipeChoicesfm)
'Maggie Wolf
'The purpose of this form is to display the all of the recipe choices that the users have.
'Displays the hotdish recipe.
Private Sub ChickenHotdishgo_Click()
    Firstfm.Hide
    RecipeChoicesfm.Hide
    ChickenHotdishfm.Show
    Lasagnafm.Hide
    TacoSaladfm.Hide
    Fajitasfm.Hide
    Mushfm.Hide
    SloppyJoesfm.Hide
    ChickenStuffingfm.Hide
    StirFryfm.Hide
End Sub
'Displays the Chicken and Stuffing recipe.
Private Sub chickenstuffinggo_Click()
    Firstfm.Hide
    RecipeChoicesfm.Hide
    ChickenHotdishfm.Hide
    Lasagnafm.Hide
    TacoSaladfm.Hide
    Fajitasfm.Hide
    Mushfm.Hide
    SloppyJoesfm.Hide
    ChickenStuffingfm.Show
    StirFryfm.Hide
End Sub
'Displays the Fajitas recipe.
Private Sub Fajitasgo_Click()
    Firstfm.Hide
    RecipeChoicesfm.Hide
    ChickenHotdishfm.Hide
    Lasagnafm.Hide
    TacoSaladfm.Hide
    Fajitasfm.Show
    Mushfm.Hide
    SloppyJoesfm.Hide
    ChickenStuffingfm.Hide
    StirFryfm.Hide
End Sub
'Displays the introduction form.
Private Sub firstgo_Click()
    Firstfm.Show
    RecipeChoicesfm.Hide
    ChickenHotdishfm.Hide
    Lasagnafm.Hide
    TacoSaladfm.Hide
    Fajitasfm.Hide
    Mushfm.Hide
    SloppyJoesfm.Hide
    ChickenStuffingfm.Hide
    StirFryfm.Hide
End Sub
'Displays the lasagna recipe.
Private Sub lasagnago_Click()
    Firstfm.Hide
    RecipeChoicesfm.Hide
    ChickenHotdishfm.Hide
    Lasagnafm.Show
    TacoSaladfm.Hide
    Fajitasfm.Hide
    Mushfm.Hide
    SloppyJoesfm.Hide
    ChickenStuffingfm.Hide
    StirFryfm.Hide
End Sub
'Displays the mush recipe.
Private Sub mushgo_Click()
    Firstfm.Hide
    RecipeChoicesfm.Hide
    ChickenHotdishfm.Hide
    Lasagnafm.Hide
    TacoSaladfm.Hide
    Fajitasfm.Hide
    Mushfm.Show
    SloppyJoesfm.Hide
    ChickenStuffingfm.Hide
    StirFryfm.Hide
End Sub
'Ends the program.
Private Sub Quit_Click()
    End
End Sub
'Displays the Sloppy Joe recipe.
Private Sub SloppyJoesgo_Click()
    Firstfm.Hide
    RecipeChoicesfm.Hide
    ChickenHotdishfm.Hide
    Lasagnafm.Hide
    TacoSaladfm.Hide
    Fajitasfm.Hide
    Mushfm.Hide
    SloppyJoesfm.Show
    ChickenStuffingfm.Hide
    StirFryfm.Hide
End Sub
'Displays the Stir Fry recipe.
Private Sub StirFrygo_Click()
    Firstfm.Hide
    RecipeChoicesfm.Hide
    ChickenHotdishfm.Hide
    Lasagnafm.Hide
    TacoSaladfm.Hide
    Fajitasfm.Hide
    Mushfm.Hide
    SloppyJoesfm.Hide
    ChickenStuffingfm.Hide
    StirFryfm.Show
End Sub
'Displays the Taco Salad recipe.
Private Sub TacoSaladgo_Click()
    Firstfm.Hide
    RecipeChoicesfm.Hide
    ChickenHotdishfm.Hide
    Lasagnafm.Hide
    TacoSaladfm.Show
    Fajitasfm.Hide
    Mushfm.Hide
    SloppyJoesfm.Hide
    ChickenStuffingfm.Hide
    StirFryfm.Hide
End Sub
