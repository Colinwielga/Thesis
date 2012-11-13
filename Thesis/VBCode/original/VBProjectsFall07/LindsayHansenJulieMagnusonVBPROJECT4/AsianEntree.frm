VERSION 5.00
Begin VB.Form AsianEntree 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Form1"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FF00FF&
      Height          =   1575
      Left            =   600
      ScaleHeight     =   1515
      ScaleWidth      =   1995
      TabIndex        =   8
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton cmdSeePicAE 
      BackColor       =   &H00C000C0&
      Caption         =   "Show Me the Picture please!"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CommandButton cmdShowIn4 
      BackColor       =   &H00C000C0&
      Caption         =   "Show Me the Ingredients please!"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton cmdShowDir4 
      BackColor       =   &H00C000C0&
      Caption         =   "Show Me the Directions please!"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   2895
   End
   Begin VB.CommandButton cmdSwitch7 
      BackColor       =   &H00C000C0&
      Caption         =   "Back to Main Form"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   2295
   End
   Begin VB.PictureBox picAsianEntree 
      BackColor       =   &H00FF00FF&
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   3600
      ScaleHeight     =   2835
      ScaleWidth      =   6915
      TabIndex        =   2
      Top             =   1200
      Width           =   6975
   End
   Begin VB.CommandButton cmdSortAlph4 
      BackColor       =   &H00C000C0&
      Caption         =   "Sort in Alphabetical Order (make sure ingredients is open)"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   2775
   End
   Begin VB.CommandButton cmdSearchIn4 
      BackColor       =   &H00C000C0&
      Caption         =   "Search for Ingredients"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Label lblAsianEntreeTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      Caption         =   "Chicken Katsu with Tonkatsu Sauce"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      TabIndex        =   6
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "AsianEntree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ctr As Integer

Private Sub cmdSearchIn4_Click()
Dim found As Boolean
Dim I As Integer
Dim NameOfIngredient As String

NameOfIngredient = InputBox("Enter an ingredient, make sure your spelling is correct!", "Ingredient Search") 'this creates an input box for user to input information
picAsianEntree.Cls  'clears the form
I = 0
found = False

Do While ((Not found) And (I < ctr))    'this loop searches for the ingredient that the user is looking for
    I = I + 1
    If NameOfIngredient = ingr(I) Then found = True
Loop

If (Not found) Then     'this if statement prints out the ingredient or tells the user the ingredient is not in the recipe
    picAsianEntree.Print NameOfIngredient; " was not in recipe!"
    Else
    picAsianEntree.Print NameOfIngredient
End If

End Sub

Private Sub cmdSeePicAE_Click()
    Picture3.Picture = LoadPicture(App.Path & "\AsianEntreePic.jpg")     'this button when clicked loads a picture
End Sub

Private Sub cmdShowDir4_Click()
Dim directions(1 To 20) As String

picAsianEntree.Cls  'clears the form

Open App.Path & "\AsianEntreeDirections.txt" For Input As #8     'this opens a file and inputs the directions in the form
    Do Until EOF(8)
    ctr = ctr + 1
        Input #8, directions(ctr)
        picAsianEntree.Print directions(ctr)
    Loop
    Close #8

End Sub

Private Sub cmdShowIn4_Click()

picAsianEntree.Cls  'clears the form
ctr = 0
Open App.Path & "\AsianEntreeIngredients.txt" For Input As #7    'this opens a file and inputs ingredients and amount in the form
    Do Until EOF(7)
    ctr = ctr + 1
        Input #7, amt(ctr), ingr(ctr)
        picAsianEntree.Print amt(ctr), ingr(ctr)
    Loop
    Close #7
End Sub

Private Sub cmdSortAlph4_Click()
Dim pass, temp2, I As Integer
Dim temp As String
Dim comp As Integer

picAsianEntree.Cls  'clears the form

For pass = 1 To ctr - 1     'this puts the ingredients in alphabetical order
For comp = 1 To ctr - pass
    If ingr(comp) > ingr(comp + 1) Then
        temp = ingr(comp)
        temp2 = amt(comp)
        ingr(comp) = ingr(comp + 1)
        amt(comp) = amt(comp + 1)
        ingr(comp + 1) = temp
        amt(comp + 1) = temp2
    End If
Next comp
Next pass

For I = 1 To ctr
    picAsianEntree.Print amt(I), ingr(I)     'this prints the amount and the ingredient
Next I
End Sub

Private Sub cmdSwitch7_Click()
Main.Show   'this shows the main form
AsianEntree.Hide    'this hides the asian entree form
End Sub
