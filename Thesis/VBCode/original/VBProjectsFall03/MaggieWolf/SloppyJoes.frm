VERSION 5.00
Begin VB.Form SloppyJoesfm 
   BackColor       =   &H00000080&
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14250
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   14250
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   4935
      Left            =   6000
      Picture         =   "SloppyJoes.frx":0000
      ScaleHeight     =   4875
      ScaleWidth      =   4155
      TabIndex        =   4
      Top             =   2400
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   1680
      TabIndex        =   3
      Top             =   6360
      Width           =   1995
   End
   Begin VB.PictureBox Picture2 
      Height          =   1935
      Left            =   1800
      Picture         =   "SloppyJoes.frx":1E9B
      ScaleHeight     =   1875
      ScaleWidth      =   1875
      TabIndex        =   2
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton ORecipes5 
      Caption         =   "WHat other recipes can I make?"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1320
      TabIndex        =   1
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
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
      Left            =   11280
      TabIndex        =   5
      Top             =   10920
      Width           =   1695
   End
   Begin VB.Label Sloppy 
      BackColor       =   &H00000080&
      Caption         =   "Sloppy Joes"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2880
      TabIndex        =   0
      Top             =   360
      Width           =   5055
   End
End
Attribute VB_Name = "SloppyJoesfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sloppy Joes(SloppyJoesfm)
'Maggie Wolf
'The purpose of this form is to display the recipe.
'This button ends the program.
Private Sub Command1_Click()
    End
End Sub

Private Sub ORecipes5_Click()
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
