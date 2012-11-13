VERSION 5.00
Begin VB.Form AsianDessert 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form1"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0000FFFF&
      Height          =   1335
      Left            =   480
      ScaleHeight     =   1275
      ScaleWidth      =   1755
      TabIndex        =   8
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton cmdSeePicAD 
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
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton cmdShowIn6 
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
      TabIndex        =   5
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton cmdShowDir6 
      BackColor       =   &H0080FFFF&
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   2655
   End
   Begin VB.CommandButton cmdSwitch9 
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Width           =   2175
   End
   Begin VB.PictureBox picAsianDessert 
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
      Height          =   3135
      Left            =   3000
      ScaleHeight     =   3075
      ScaleWidth      =   7755
      TabIndex        =   2
      Top             =   840
      Width           =   7815
   End
   Begin VB.CommandButton cmdSortAlph6 
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
      Height          =   1215
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   2775
   End
   Begin VB.CommandButton cmdSearchIn6 
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
      Height          =   1215
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label lblAsianDessertTitle 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Chi Chi Dango Mochi"
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
      Left            =   3960
      TabIndex        =   6
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "AsianDessert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ctr As Integer


Private Sub cmdSearchIn6_Click()
Dim found As Boolean
Dim I As Integer
Dim NameOfIngredient As String

NameOfIngredient = InputBox("Enter an ingredient, make sure your spelling is correct!", "Ingredient Search")    'this creates an input box for user to input information
picAsianDessert.Cls 'clears the form
I = 0
found = False

Do While ((Not found) And (I < ctr))    'this loop searches for the ingredient that the user is looking for
    I = I + 1
    If NameOfIngredient = ingr(I) Then found = True
Loop

If (Not found) Then     'this if statement prints out the ingredient or tells the user the ingredient is not in the recipe
    picAsianDessert.Print NameOfIngredient; " was not in recipe!"
    Else
    picAsianDessert.Print NameOfIngredient
End If
End Sub

Private Sub cmdSeePicAD_Click()
    Picture2.Picture = LoadPicture(App.Path & "\AsianDessertPic.jpg")   'this button when clicked loads a picture
End Sub

Private Sub cmdShowIn6_Click()

picAsianDessert.Cls 'this clears the form
ctr = 0
Open App.Path & "\AsianDessertIngredients.txt" For Input As #11     'this opens a file and inputs ingredients and amount in the form
    Do Until EOF(11)
    ctr = ctr + 1
        Input #11, amt(ctr), ingr(ctr)
        picAsianDessert.Print amt(ctr), ingr(ctr)
    Loop
    Close #11
End Sub
Private Sub cmdShowDir6_Click()
Dim directions(1 To 20) As String

picAsianDessert.Cls     'this clears the form

Open App.Path & "\AsianDessertDirections.txt" For Input As #12  'this opens a file and inputs the directions in the form
    Do Until EOF(12)
    ctr = ctr + 1
        Input #12, directions(ctr)
        picAsianDessert.Print directions(ctr)
    Loop
    Close #12

End Sub


Private Sub cmdSortAlph6_Click()
Dim pass, temp2, I As Integer
Dim temp As String
Dim comp As Integer

picAsianDessert.Cls     'this clears the form

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
    picAsianDessert.Print amt(I), ingr(I)   'this prints the amount and the ingredient
Next I
End Sub

Private Sub cmdSwitch9_Click()
Main.Show       'this shows the main form
AsianDessert.Hide 'this hides the asian dessert form
End Sub
