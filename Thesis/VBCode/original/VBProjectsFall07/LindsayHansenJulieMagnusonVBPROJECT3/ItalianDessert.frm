VERSION 5.00
Begin VB.Form ItalianDessert 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form1"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   10275
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture5 
      BackColor       =   &H0000FFFF&
      Height          =   1575
      Left            =   480
      ScaleHeight     =   1515
      ScaleWidth      =   2115
      TabIndex        =   8
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton cmdSeePicID 
      BackColor       =   &H0080FFFF&
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
      Height          =   975
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton cmdShowDir5 
      BackColor       =   &H0080FFFF&
      Caption         =   "Show me the Directions please!"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   3015
   End
   Begin VB.CommandButton cmdShowIn5 
      BackColor       =   &H0080FFFF&
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
      Height          =   1335
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton cmdSwitch8 
      BackColor       =   &H0080FFFF&
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
      Height          =   615
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5520
      Width           =   2415
   End
   Begin VB.PictureBox picItalianDessert 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   3720
      ScaleHeight     =   2715
      ScaleWidth      =   6075
      TabIndex        =   2
      Top             =   1080
      Width           =   6135
   End
   Begin VB.CommandButton cmdSortAlph5 
      BackColor       =   &H0080FFFF&
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   2895
   End
   Begin VB.CommandButton cmdSearchIn5 
      BackColor       =   &H0080FFFF&
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Label lblItalianDessertTitle 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Tiramisu"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   6
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "ItalianDessert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ctr As Integer

Private Sub cmdSearchIn5_Click()
Dim found As Boolean
Dim I As Integer
Dim NameOfIngredient As String

NameOfIngredient = InputBox("Enter an ingredient, make sure your spelling is correct!", "Ingredient Search")    'this creates an input box for user to input information
picItalianDessert.Cls   'clears the form
I = 0
found = False

Do While ((Not found) And (I < ctr))     'this loop searches for the ingredient that the user is looking for
    I = I + 1
    If NameOfIngredient = ingr(I) Then found = True
Loop

If (Not found) Then     'this if statement prints out the ingredient or tells the user the ingredient is not in the recipe
    picItalianDessert.Print NameOfIngredient; " was not in recipe!"
    Else
    picItalianDessert.Print NameOfIngredient
End If
End Sub

Private Sub cmdSeePicID_Click()
    Picture5.Picture = LoadPicture(App.Path & "\ItalianDessertPic.jpg") 'this button when clicked loads a picture
End Sub

Private Sub cmdShowIn5_Click()
picItalianDessert.Cls   'clears the form
ctr = 0
Open App.Path & "\ItalianDessertIngredients.txt" For Input As #9     'this opens a file and inputs ingredients and amount in the form
    Do Until EOF(9)
    ctr = ctr + 1
        Input #9, amt(ctr), ingr(ctr)
        picItalianDessert.Print amt(ctr), ingr(ctr)
    Loop
    Close #9
End Sub
Private Sub cmdShowDir5_Click()
Dim directions(1 To 20) As String

picItalianDessert.Cls   'clears the form

Open App.Path & "\ItalianDessertDirections.txt" For Input As #10   'this opens a file and inputs the directions in the form
    Do Until EOF(10)
    ctr = ctr + 1
        Input #10, directions(ctr)
        picItalianDessert.Print directions(ctr)
    Loop
    Close #10

End Sub

Private Sub cmdSortAlph5_Click()
Dim pass, temp2, I As Integer
Dim temp As String
Dim comp As Integer

picItalianDessert.Cls   'clears the form

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
    picItalianDessert.Print amt(I), ingr(I) 'this prints the amount and the ingredient
Next I
End Sub

Private Sub cmdSwitch8_Click()
Main.Show       'this shows the main form
ItalianDessert.Hide     'this hides the italian dessert form
End Sub
