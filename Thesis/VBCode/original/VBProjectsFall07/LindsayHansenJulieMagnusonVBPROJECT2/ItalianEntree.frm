VERSION 5.00
Begin VB.Form ItalianEntree 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Form1"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   10455
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00C000C0&
      Height          =   1455
      Left            =   480
      ScaleHeight     =   1395
      ScaleWidth      =   1875
      TabIndex        =   8
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton cmdSeePicIE 
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CommandButton cmdShowDir3 
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
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   2535
   End
   Begin VB.CommandButton cmdShowIn3 
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
      Height          =   975
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   2535
   End
   Begin VB.CommandButton cmdSwitch6 
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5400
      Width           =   2415
   End
   Begin VB.PictureBox picItalianEntree 
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
      Left            =   2880
      ScaleHeight     =   2835
      ScaleWidth      =   7395
      TabIndex        =   2
      Top             =   1080
      Width           =   7455
   End
   Begin VB.CommandButton cmdSortAlph3 
      BackColor       =   &H00C000C0&
      Caption         =   "Sort Ingredients in Alphabetical Order (make sure ingredients is open)"
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
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CommandButton cmdSearchIn3 
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   2655
   End
   Begin VB.Label lblItalianEntreeTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      Caption         =   "Penne All'Arrabbiata"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3600
      TabIndex        =   6
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "ItalianEntree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ctr As Integer

Private Sub cmdSearchIn3_Click()
Dim found As Boolean
Dim I As Integer
Dim NameOfIngredient As String

NameOfIngredient = InputBox("Enter an ingredient, make sure your spelling is correct!", "Ingredient Search")    'this creates an input box for user to input information
picItalianEntree.Cls    'clears the form
I = 0
found = False

Do While ((Not found) And (I < ctr))    'this loop searches for the ingredient that the user is looking for
    I = I + 1
    If NameOfIngredient = ingr(I) Then found = True
Loop

If (Not found) Then 'this if statement prints out the ingredient or tells the user the ingredient is not in the recipe
    picItalianEntree.Print NameOfIngredient; " was not in recipe!"
    Else
    picItalianEntree.Print NameOfIngredient
End If

End Sub

Private Sub cmdSeePicIE_Click()
    Picture6.Picture = LoadPicture(App.Path & "\ItalianEntreePic.jpg")  'this button when clicked loads a picture
End Sub

Private Sub cmdShowDir3_Click()
Dim directions(1 To 20) As String

picItalianEntree.Cls    'clears the form

Open App.Path & "\ItalianEntreeDirections.txt" For Input As #6  'this opens a file and inputs directions in the form
    Do Until EOF(6)
    ctr = ctr + 1
        Input #6, directions(ctr)
        picItalianEntree.Print directions(ctr)
    Loop
    Close #6

End Sub

Private Sub cmdShowIn3_Click()

picItalianEntree.Cls    'clears the form
ctr = 0
Open App.Path & "\ItalianEntreeIngredients.txt" For Input As #5 'this opens a file and inputs ingredients and amount in the form
    Do Until EOF(5)
    ctr = ctr + 1
        Input #5, amt(ctr), ingr(ctr)
        picItalianEntree.Print amt(ctr), ingr(ctr)
    Loop
    Close #5
End Sub


Private Sub cmdSortAlph3_Click()
Dim pass, temp2, I As Integer
Dim temp As String
Dim comp As Integer

picItalianEntree.Cls    'clears the form

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
    picItalianEntree.Print amt(I), ingr(I)  'this prints the amount and the ingredient
Next I

End Sub

Private Sub cmdSwitch6_Click()
Main.Show   'this shows the main form
ItalianEntree.Hide  'this hides the italian entree form
End Sub
