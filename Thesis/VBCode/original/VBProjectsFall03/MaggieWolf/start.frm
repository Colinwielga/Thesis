VERSION 5.00
Begin VB.Form Firstfm 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ORecipes 
      Caption         =   "See all Recipe Choices"
      Height          =   1095
      Left            =   8280
      TabIndex        =   11
      Top             =   7920
      Width           =   2655
   End
   Begin VB.CommandButton Quit 
      Caption         =   "Quit"
      Height          =   1095
      Left            =   480
      TabIndex        =   10
      Top             =   7920
      Width           =   2295
   End
   Begin VB.PictureBox Picture2 
      Height          =   2175
      Left            =   8520
      Picture         =   "start.frx":0000
      ScaleHeight     =   2115
      ScaleWidth      =   3195
      TabIndex        =   9
      Top             =   120
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   120
      Picture         =   "start.frx":2D73
      ScaleHeight     =   2115
      ScaleWidth      =   2475
      TabIndex        =   8
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton GetRecipe 
      Caption         =   "Get The Recipe!"
      Height          =   1935
      Left            =   3600
      TabIndex        =   7
      Top             =   7440
      Width           =   3975
   End
   Begin VB.PictureBox Results 
      Height          =   855
      Left            =   240
      ScaleHeight     =   795
      ScaleWidth      =   11115
      TabIndex        =   6
      Top             =   6240
      Width           =   11175
   End
   Begin VB.TextBox time 
      Height          =   735
      Left            =   7200
      TabIndex        =   5
      Top             =   3720
      Width           =   4455
   End
   Begin VB.TextBox meats 
      Height          =   615
      Left            =   7200
      TabIndex        =   3
      Top             =   2640
      Width           =   4455
   End
   Begin VB.CommandButton recipechoices 
      Caption         =   "What should I make for dinner?"
      Height          =   855
      Left            =   2520
      TabIndex        =   0
      Top             =   4920
      Width           =   6375
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Maggie Wolf"
      Height          =   375
      Left            =   9720
      TabIndex        =   12
      Top             =   10560
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "How much time do you have to prepare the meal?   (in minutes)"
      Height          =   855
      Left            =   360
      TabIndex        =   4
      Top             =   3720
      Width           =   5655
   End
   Begin VB.Label CB 
      BackColor       =   &H000000FF&
      Caption         =   "Would you like to eat chicken or beef?"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2760
      Width           =   4335
   End
   Begin VB.Label chickbeef 
      BackColor       =   &H000000FF&
      Caption         =   "Chicken and Beef Entrees"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      TabIndex        =   1
      Top             =   1200
      Width           =   5655
   End
End
Attribute VB_Name = "Firstfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Recipe Search by Maggie Wolf
'Chicken and Beef Entrees (Firstfm)
'The purpose of this form is to find out what kind of meat the user wants to eat and also find out how much time they have to make it.
'This form then directs the user to the recipe that they should make.
Option Explicit
Public PATH As String
'Dim Variables
Dim Entree(1 To 8) As String, meat(1 To 8) As String, Minutes(1 To 8) As Single, I As Integer, G As Integer

'Show first form only
Private Sub ORecipes_Click()
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
'End program
Private Sub Quit_Click()
    End
End Sub
'Open file with data and sort into an array.
Private Sub recipechoices_Click()
PATH = "N:\CS130\handin\Maggie Wolf\"
Open PATH & "recipes.txt" For Input As #1
For G = 1 To 8
    Input #1, Entree(G), meat(G), Minutes(G)
Next G
Close #1
Dim Found As Boolean, NotFound As Boolean
Dim T As Integer, M As String
T = time.Text
M = meats.Text
I = 0
'Find where data from textboxes fits in with the data in the array.
NotFound = True
    Do While NotFound And I < 8
        I = I + 1
        If T >= Minutes(I) Then
            If LCase(M) = meat(I) Then
                NotFound = False
            End If
        End If
    Loop
'Print message box or message, depending on where the variables fit.
If NotFound Then
        MsgBox "Sorry, you'll have to make a peanut butter and jelly sandwich, you don't have time for anything else.", , "yum"
    Else
        Results.Print "You should make "; Entree(I); " tonight because it is made of "; meat(I); " and it takes "; Minutes(I); " minutes to prepare."
End If
End Sub
'Sort through data and find what recipe fits.
Private Sub GetRecipe_Click()
Open PATH & "recipes.txt" For Input As #1
For G = 1 To 8
    Input #1, Entree(G), meat(G), Minutes(G)
Next G
Close #1
Dim Found As Boolean, NotFound As Boolean
Dim T As Integer, M As String
T = time.Text
M = meats.Text
I = 0
NotFound = True
    Do While NotFound And I < 8
        I = I + 1
        If T >= Minutes(I) Then
            If LCase(M) = meat(I) Then
                NotFound = False
            End If
        End If
    Loop
'If recipe doesn't fit, than message box is displayed.
If NotFound Then
        MsgBox "Put peanut butter on one side of bread, jelly on the other.", , "yum"
    Else
'Select Case sorts through data to find which form should be displayed.
        If meat(I) = meat(1) Then
            Select Case Minutes(I)
                Case Is >= 70
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
                Case 20 To 69
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
                Case 15 To 19
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
            End Select
        End If
        If meat(I) = meat(2) Then
            Select Case Minutes(I)
                Case Is >= 60
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
                Case 45 To 59
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
                Case 30 To 44
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
                Case 25 To 29
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
                Case 20 To 24
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
                End Select
            End If
End If
End Sub
