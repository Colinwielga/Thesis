VERSION 5.00
Begin VB.Form ChickenStuffingfm 
   BackColor       =   &H00008080&
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13785
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   13785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Quit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      TabIndex        =   4
      Top             =   6960
      Width           =   2415
   End
   Begin VB.CommandButton ORecipes6 
      Caption         =   "What other recipes can I make?"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   480
      TabIndex        =   3
      Top             =   4920
      Width           =   4095
   End
   Begin VB.PictureBox Picture2 
      Height          =   2535
      Left            =   840
      Picture         =   "ChickenStuffing.frx":0000
      ScaleHeight     =   2475
      ScaleWidth      =   3435
      TabIndex        =   2
      Top             =   2040
      Width           =   3495
   End
   Begin VB.PictureBox Picture1 
      Height          =   4095
      Left            =   6120
      Picture         =   "ChickenStuffing.frx":3C4A
      ScaleHeight     =   4035
      ScaleWidth      =   4155
      TabIndex        =   1
      Top             =   2760
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00008080&
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
      Left            =   10920
      TabIndex        =   5
      Top             =   10800
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008080&
      Caption         =   "Chicken and Stuffing"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2520
      TabIndex        =   0
      Top             =   600
      Width           =   5895
   End
End
Attribute VB_Name = "ChickenStuffingfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Chicken and Stuffing (ChickenStuffingfm)
'Maggie Wolf
'The purpose of this form is to display the recipe.
'This button ends the program.
Private Sub ORecipes6_Click()
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
