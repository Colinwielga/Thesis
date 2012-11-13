VERSION 5.00
Begin VB.Form AsianApp 
   BackColor       =   &H00C0C000&
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   10230
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picture1 
      BackColor       =   &H00C0FFC0&
      Height          =   1455
      Left            =   840
      ScaleHeight     =   1395
      ScaleWidth      =   1755
      TabIndex        =   8
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmdSeePicAA 
      BackColor       =   &H00FF8080&
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton cmdShowDir2 
      BackColor       =   &H00FF8080&
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   2775
   End
   Begin VB.CommandButton cmdShowIn2 
      BackColor       =   &H00FF8080&
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton cmdSwitch5 
      BackColor       =   &H00FF8080&
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5280
      Width           =   2415
   End
   Begin VB.PictureBox picAsianApp 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   3480
      ScaleHeight     =   2475
      ScaleWidth      =   6555
      TabIndex        =   2
      Top             =   1200
      Width           =   6615
   End
   Begin VB.CommandButton cmdSortAlph2 
      BackColor       =   &H00FF8080&
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      Width           =   2895
   End
   Begin VB.CommandButton cmdSearchIn2 
      BackColor       =   &H00FF8080&
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3840
      Width           =   2895
   End
   Begin VB.Label lblAsianAppTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Marinated Tofu"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      TabIndex        =   6
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "AsianApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ctr As Integer

Private Sub cmdSearchIn2_Click()
Dim found As Boolean
Dim I As Integer
Dim NameOfIngredient As String

NameOfIngredient = InputBox("Enter an ingredient, make sure your spelling is correct!", "Ingredient Search")    'this creates an input box for user to input information
picAsianApp.Cls     'clears the form
I = 0
found = False
        
Do While ((Not found) And (I < ctr))    'this loop searches for the ingredient that the user is looking for
    I = I + 1
    If NameOfIngredient = ingr(I) Then found = True
Loop

If (Not found) Then     'this if statement prints out the ingredient or tells the user the ingredient is not in the recipe
    picAsianApp.Print NameOfIngredient; " was not in recipe!"
    Else
    picAsianApp.Print NameOfIngredient
End If


End Sub

Private Sub cmdSeePicAA_Click()
    picture1.Picture = LoadPicture(App.Path & "\AsianAppPic.jpg")   'this button when clicked loads a picture
End Sub

Private Sub cmdShowDir2_Click()
Dim directions(1 To 20) As String

picAsianApp.Cls     'this clears the form

Open App.Path & "\AsianAppDirections.txt" For Input As #4   'this opens a file and inputs the directions in the form
    Do Until EOF(4)
    ctr = ctr + 1
        Input #4, directions(ctr)
        picAsianApp.Print directions(ctr)
    Loop
    Close #4

End Sub

Private Sub cmdShowIn2_Click()

picAsianApp.Cls     'this clears the form
ctr = 0
Open App.Path & "\AsianAppIngredients.txt" For Input As #3  'this opens a file and inputs ingredients and amount in the form
    Do Until EOF(3)
    ctr = ctr + 1
        Input #3, amt(ctr), ingr(ctr)
        picAsianApp.Print amt(ctr), ingr(ctr)
    Loop
    Close #3
End Sub

Private Sub cmdSortAlph2_Click()

Dim pass, temp2, I As Integer
Dim temp As String
Dim comp As Integer

picAsianApp.Cls     'this clears the form

For pass = 1 To ctr - 1         'this puts the ingredients in alphabetical order
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
    picAsianApp.Print amt(I), ingr(I)   'this prints the amount and the ingredient
Next I

End Sub

Private Sub cmdSwitch5_Click()
Main.Show       'this shows the main form
AsianApp.Hide   'this hides the asian appetizer form
End Sub


