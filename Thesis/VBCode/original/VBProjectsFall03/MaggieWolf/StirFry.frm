VERSION 5.00
Begin VB.Form StirFryfm 
   BackColor       =   &H000080FF&
   Caption         =   "Form1"
   ClientHeight    =   10590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13320
   LinkTopic       =   "Form1"
   ScaleHeight     =   10590
   ScaleWidth      =   13320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Quit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1680
      TabIndex        =   4
      Top             =   6600
      Width           =   2655
   End
   Begin VB.CommandButton ORecipes7 
      Caption         =   "What Other Recipes Can I Make?"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1080
      TabIndex        =   3
      Top             =   4560
      Width           =   3975
   End
   Begin VB.PictureBox Picture2 
      Height          =   2535
      Left            =   1200
      Picture         =   "StirFry.frx":0000
      ScaleHeight     =   2475
      ScaleWidth      =   3555
      TabIndex        =   2
      Top             =   1800
      Width           =   3615
   End
   Begin VB.PictureBox Picture1 
      Height          =   5895
      Left            =   5880
      Picture         =   "StirFry.frx":40F2
      ScaleHeight     =   5835
      ScaleWidth      =   4035
      TabIndex        =   1
      Top             =   1680
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackColor       =   &H000080FF&
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
      Top             =   9960
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "Chicken Stir Fry"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2640
      TabIndex        =   0
      Top             =   480
      Width           =   5295
   End
End
Attribute VB_Name = "StirFryfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Stir Fry (StirFryfm)
'Maggie Wolf
'The purpose of this form is to display the recipe.
'This button ends the program.
Private Sub ORecipes7_Click()
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
