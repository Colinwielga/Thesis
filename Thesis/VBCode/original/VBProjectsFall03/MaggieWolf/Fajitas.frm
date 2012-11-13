VERSION 5.00
Begin VB.Form Fajitasfm 
   BackColor       =   &H00FF00FF&
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13005
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   13005
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   1935
      Left            =   1800
      Picture         =   "Fajitas.frx":0000
      ScaleHeight     =   1875
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton ORecipes3 
      Caption         =   "What other recipes are there?"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   960
      TabIndex        =   3
      Top             =   4920
      Width           =   2895
   End
   Begin VB.CommandButton Quit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1320
      TabIndex        =   2
      Top             =   6960
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   5775
      Left            =   5280
      Picture         =   "Fajitas.frx":1545
      ScaleHeight     =   5715
      ScaleWidth      =   4755
      TabIndex        =   1
      Top             =   2400
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF00FF&
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
      Left            =   10440
      TabIndex        =   5
      Top             =   10680
      Width           =   1695
   End
   Begin VB.Label Fajitas 
      BackColor       =   &H00FF00FF&
      Caption         =   "Fajitas"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3240
      TabIndex        =   0
      Top             =   480
      Width           =   3615
   End
End
Attribute VB_Name = "Fajitasfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Fajitas (Fajitasfm)
'Maggie Wolf
'The purpose of this form is to display the recipe.
'This button ends the program.
Private Sub ORecipes3_Click()
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
